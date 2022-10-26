using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using tff.main.Models;
using Application = Microsoft.Office.Interop.Word.Application;
using MessageBox = System.Windows.Forms.MessageBox;
using Range = Microsoft.Office.Interop.Word.Range;

namespace tff.main.Handlers;

public static class DocProcessHandler
{
    private static BackgroundWorker _worker;
    private static Entry _entry;
    private static readonly Regex _etalonRequestPattern = new("#a/\\d{1,}/request");
    private static readonly Regex _etalonResponsePattern = new("#a/\\d{1,}/response");
    private static readonly Regex _testPattern = new("#b/\\d{1,}");
    private static readonly Regex _tagPattern = new("<\\w*?((?:.(?!\\1|>))*.?)\\1?>");

    public static bool Execute(Entry entry)
    {
        InitWorker(entry);
        _worker.RunWorkerAsync();

        return true;
    }

    private static void ExecuteInternal(DoWorkEventArgs e)
    {
        var application = new Application();
        Document newDoc = null;
        var currentStep = 0;

        try
        {
            var fileInfo = new FileInfo(_entry.TargetFile ?? throw new InvalidOperationException());
            var newFilePath = $"{_entry.SavePath}\\{fileInfo.Name.Split(".")[0]}_test.{fileInfo.Extension}";

            if (fileInfo.Exists)
            {
                fileInfo.CopyTo(newFilePath, true);
            }

            newDoc = application.Documents.Open(newFilePath);

            var fullDocText = newDoc.Range()
                .Text;

            _entry.TotalCount = _etalonRequestPattern.Matches(fullDocText)
                                    .Count
                                + _etalonResponsePattern.Matches(fullDocText)
                                    .Count
                                + _testPattern.Matches(fullDocText)
                                    .Count;

            if (_entry.TotalCount == 0)
            {
                throw new Exception(
                    "Нет шаблонов для заполнения. Проверьте исходный файл на существование шаблонов типа #a/1/method или #b/1"
                );
            }

            for (var i = 1; i <= newDoc.Sections.Count; i++)
            {
                if (_worker.CancellationPending)
                {
                    e.Cancel = true;

                    break;
                }

                var wordRange = newDoc.Sections[i]
                    .Range;

                SetXmlText(wordRange,ref currentStep);
            }

            newDoc.Save();
        }
        finally
        {
            newDoc?.Close();
            application.Quit();
        }
    }

    private static void SetXmlText(Range wordRange, ref int currentStep)
    {
        var requestMatches = _etalonRequestPattern.Matches(wordRange.Text)
            .OrderBy(x => x.Index)
            .ToArray();

        var responseMatches = _etalonResponsePattern.Matches(wordRange.Text)
            .OrderBy(x => x.Index)
            .ToArray();

        var testMatches = _testPattern.Matches(wordRange.Text)
            .OrderBy(x => x.Index)
            .ToArray();

        if (requestMatches.Length == 0 && responseMatches.Length == 0 && testMatches.Length == 0)
        {
            return;
        }

        var iterateCount = Math.Max(Math.Max(requestMatches.Length, responseMatches.Length), testMatches.Length);
        for (var i = 0; i < iterateCount; i++)
        {
            if (_worker.CancellationPending)
            {
                return;
            }

            var currentMatch = requestMatches.Length > i ? requestMatches[i] : null;

            if (currentMatch != null)
            {
                SetXmlRangeParams(currentMatch, wordRange, ProcessType.EtalonRequest);
                _entry.TotalProgress = GetTotalPercent(++currentStep);
            }

            currentMatch = responseMatches.Length > i ? responseMatches[i] : null;

            if (currentMatch != null)
            {
                SetXmlRangeParams(currentMatch, wordRange, ProcessType.EtalonResponse);
                _entry.TotalProgress = GetTotalPercent(++currentStep);
            }

            currentMatch = testMatches.Length > i ? testMatches[i] : null;

            if (currentMatch != null)
            {
                SetXmlRangeParams(currentMatch, wordRange, ProcessType.Test);
                _entry.TotalProgress = GetTotalPercent(++currentStep);
            }
        }
    }

    private static string GetXmlPath(string xmlNumber, ProcessType processType)
    {
        if (string.IsNullOrEmpty(xmlNumber))
            return string.Empty;

        return processType switch
        {
            ProcessType.EtalonRequest => $"{_entry.EtalonFolder}/{xmlNumber}/Request.xml",
            ProcessType.EtalonResponse => $"{_entry.EtalonFolder}/{xmlNumber}/Response.xml",
            ProcessType.Test => $"{_entry.TestFolder}/{xmlNumber}.xsl",
            _ => string.Empty,
        };
    }

    private static void InitWorker(Entry entry)
    {
        _worker = new BackgroundWorker();
        _worker.WorkerReportsProgress = true;
        _worker.DoWork += worker_DoWork;
        _worker.ProgressChanged += worker_ProgressChanged;
        _worker.RunWorkerCompleted += _worker_RunWorkerCompleted;
        _worker.WorkerSupportsCancellation = true;

        _entry = entry;
    }

    private static void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        if (_entry != null)
        {
            _entry.StartVisible = Visibility.Visible;
            _entry.Progress = 0;
            _entry.TotalCount = 0;
            _entry.TotalProgress = 0;
        }

        if (e.Error != null)
        {
            MessageBox.Show(e.Error.Message, "Ошибка");

            return;
        }

        MessageBox.Show($"Задача {(!e.Cancelled ? "выполнена" : "отменена")}", "Завершение");
    }

    private static void worker_DoWork(object sender, DoWorkEventArgs e)
    {
        if (_entry == null)
        {
            return;
        }

        _entry.StartVisible = Visibility.Collapsed;
        ExecuteInternal(e);
    }

    private static void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
        if (_entry != null)
        {
            _entry.Progress = e.ProgressPercentage;
        }
    }

    public static void worker_Stop()
    {
        if (_worker is {WorkerSupportsCancellation: true, CancellationPending: false,})
        {
            _worker.CancelAsync();
        }
    }

    private static int GetPercent(int current, int totalInTemplate)
    {
        var experssion = (int) decimal.Round(current / (decimal)totalInTemplate * 100);

        return experssion;
    }

    private static int GetTotalPercent(int current)
    {
        var experssion = (int)decimal.Round(current / (decimal)_entry.TotalCount * 100);

        return experssion;
    }

    private static void SetStartEndIndex(Match match, Range range, bool isNeedSetStart=true)
    {
        if(isNeedSetStart)
            range.Start += match.Index;

        range.End = range.Start + match.Length;
    }

    private static void SetXmlRangeParams(Match match, Range wordRange, ProcessType processType)
    {
        _worker.ReportProgress(0);

        var currentStep = 1;
        _entry.CurrentTemplate = match.Value;

        if (_worker.CancellationPending)
        {
            return;
        }

        var xmlPath = GetXmlPath(match.Value.Split("/")[1], processType);
        var xmlDocument = new XmlDocument();
        xmlDocument.Load(xmlPath);
        var element = XElement.Parse(xmlDocument.OuterXml);

        var settings = new XmlWriterSettings
        {
            OmitXmlDeclaration = true,
            Indent = true,
            NewLineOnAttributes = false,
        };

        var stringBuilder = new StringBuilder();

        using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
        {
            element.Save(xmlWriter);
        }

        var currentRange = wordRange.Duplicate;
        SetStartEndIndex(match, currentRange, false);

        currentRange.Text = stringBuilder.ToString();
        currentRange.Paragraphs.LineSpacing = 12;
        currentRange.Paragraphs.SpaceAfter = 0;

        wordRange.Start += stringBuilder.Length;
        stringBuilder.Clear();

        var tags = _tagPattern.Matches(currentRange.Text);
        var tagsCount = tags.Count;
        var progress = GetPercent(currentStep, tagsCount + 1);
        _worker.ReportProgress(progress);

        _entry.CurrentTemplate = $"Обработка тэгов {match.Value}";

        foreach (Match tag in tags)
        {
            if (_worker.CancellationPending)
            {
                return;
            }

            ++currentStep;
            var tagRange = currentRange.Duplicate;
            SetStartEndIndex(tag, tagRange);
            tagRange.Font.Color = WdColor.wdColorLightBlue;

            progress = GetPercent(currentStep, tagsCount+1);
            _worker.ReportProgress(progress);
        }
    }

    private enum ProcessType
    {
        EtalonRequest,
        EtalonResponse,
        Test,
    }
}