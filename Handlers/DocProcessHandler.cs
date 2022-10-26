using System;
using System.ComponentModel;
using System.IO;
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

public class DocProcessHandler
{
    private static BackgroundWorker? _worker;
    private static Entry? _entry;

    public static bool Execute(Entry entry)
    {
        InitWorker(entry);
        _worker?.RunWorkerAsync();

        return true;
    }

    private static void ExecuteInternal()
    {
        var result = new ResultFile();
        var fileInfo = new FileInfo(_entry.TargetFile);
        var newFilePath = $"{_entry.SavePath}\\{fileInfo.Name.Split(".")[0]}_test.{fileInfo.Extension}";
        var application = new Application();
        Document? newDoc = null;
        var currentStep = 0;

        try
        {
            if (fileInfo.Exists)
            {
                fileInfo.CopyTo(newFilePath, true);
            }

            newDoc = application.Documents.Open(newFilePath);

            var etalonRequestPattern = new Regex("#a/\\d{1,}/request");
            var etalonResponsePattern = new Regex("#a/\\d{1,}/response");
            var testPattern = new Regex("#b/\\d{1,}");
            var tagPattern = new Regex("<\\w*?((?:.(?!\\1|>))*.?)\\1?>");

            var fullDocText = newDoc.Range()
                .Text;

            _entry.TotalCount = etalonRequestPattern.Matches(fullDocText)
                                    .Count
                                + etalonResponsePattern.Matches(fullDocText)
                                    .Count
                                + testPattern.Matches(fullDocText)
                                    .Count
                                + tagPattern.Matches(fullDocText)
                                    .Count;

            if (_entry.TotalCount == 0)
            {
                throw new Exception(
                    "Нет шаблонов для заполнения. Проверьте исходный файл на существование шаблонов типа #a/1/method или #b/1"
                );
            }

            for (var i = 1; i <= newDoc.Sections.Count; i++)
            {
                var wordRange = newDoc.Sections[i]
                    .Range;

                var matches = etalonRequestPattern.Matches(wordRange.Text);

                SetXmlText(matches,
                    tagPattern,
                    wordRange,
                    ProcessType.EtalonRequest,
                    ref currentStep
                );

                matches = etalonResponsePattern.Matches(wordRange.Text);

                SetXmlText(matches,
                    tagPattern,
                    wordRange,
                    ProcessType.EtalonResponse,
                    ref currentStep
                );

                matches = testPattern.Matches(wordRange.Text);

                SetXmlText(matches,
                    tagPattern,
                    wordRange,
                    ProcessType.Test,
                    ref currentStep
                );
            }

            newDoc.Save();
        }
        finally
        {
            newDoc?.Close();
            application.Quit();
        }
    }

    private static void SetXmlText(MatchCollection matches,
                                   Regex tagPattern,
                                   Range wordRange,
                                   ProcessType processType,
                                   ref int currentStep)
    {
        if (matches.Count > 0)
        {
            foreach (Match match in matches)
            {
                ++currentStep;

                var stringBuilder = new StringBuilder();

                var requestNumber = match.Value.Split("/")[1];
                var currentRange = wordRange.Duplicate;

                var xmlDocument = new XmlDocument();
                var xmlPath = GetXmlPath(requestNumber, processType);
                xmlDocument.Load(xmlPath);
                var element = XElement.Parse(xmlDocument.OuterXml);

                var settings = new XmlWriterSettings
                {
                    OmitXmlDeclaration = true,
                    Indent = true,
                    NewLineOnAttributes = false,
                };

                using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
                {
                    element.Save(xmlWriter);
                }

                var findObj = currentRange.Find;
                findObj.ClearFormatting();
                findObj.Text = match.Value;
                findObj.Execute();
                currentRange.Text = stringBuilder.ToString();
                currentRange.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                currentRange.Paragraphs.LineUnitAfter = 0;
                currentRange.Paragraphs.SpaceAfter = 0;
                currentRange.Paragraphs.SpaceAfterAuto = 0;

                var progress = GetPercent(currentStep);
                _worker.ReportProgress(progress);

                var tags = tagPattern.Matches(currentRange.Text);

                foreach (Match tag in tags)
                {
                    ++currentStep;
                    var tagRange = currentRange.Duplicate;
                    tagRange.Start = tag.Index - 1;
                    var tagFind = tagRange.Find;
                    tagFind.ClearFormatting();
                    tagFind.Text = tag.Value;
                    tagFind.Execute();
                    tagRange.Font.Color = WdColor.wdColorLightBlue;

                    progress = GetPercent(currentStep);
                    _worker.ReportProgress(progress);
                }
            }
        }
    }

    private static string GetXmlPath(string xmlNumber, ProcessType processType)
    {
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

    private static void _worker_RunWorkerCompleted(object? sender, RunWorkerCompletedEventArgs e)
    {
        if (_entry != null)
        {
            _entry.StartVisible = Visibility.Visible;
            _entry.Progress = 0;
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
        if (_entry != null)
        {
            _entry.StartVisible = Visibility.Collapsed;
            ExecuteInternal();
        }
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
        if (_worker != null && _worker.WorkerSupportsCancellation && !_worker.CancellationPending)
        {
            _worker?.CancelAsync();
        }
    }

    private static int GetPercent(int current)
    {
        var experssion = (int) decimal.Round(current / (decimal) _entry.TotalCount * 100);

        return experssion;
    }

    private enum ProcessType
    {
        EtalonRequest,
        EtalonResponse,
        Test,
    }
}