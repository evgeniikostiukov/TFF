using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using tff.main.Models;
using MessageBox = System.Windows.Forms.MessageBox;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace tff.main.Handlers;

public class DocProcessHandler
{
    private readonly Regex _etalonRequestPattern;
    private readonly Regex _etalonResponsePattern;
    private readonly Regex _testPattern;
    private readonly Entry _entry;
    private BackgroundWorker _worker;

    public DocProcessHandler(Entry entry)
    {
        InitWorker();
        _etalonRequestPattern = new Regex("#a/\\d{1,}/request");
        _etalonResponsePattern = new Regex("#a/\\d{1,}/response");
        _testPattern = new Regex("#b/\\d{1,}");
        _entry = entry;
    }

    public void Execute()
    {
        _worker.RunWorkerAsync();
    }

    private void ExecuteInternal(DoWorkEventArgs e)
    {
        var fileInfo = new FileInfo(_entry.TargetFile ?? throw new InvalidOperationException());
        var newFilePath = $"{_entry.SavePath}\\{fileInfo.Name.Split(".")[0]}_ГОТОВЫЙ{fileInfo.Extension}";

        var oldDoc = WordprocessingDocument.Open(_entry.TargetFile,
            false,
            new OpenSettings
            {
                AutoSave = false,
            }
        );

        using (oldDoc.Clone(newFilePath))
        {
            oldDoc.Dispose();
        }

        using var newDoc = WordprocessingDocument.Open(newFilePath, true);

        var templateParagraphs = newDoc.MainDocumentPart?.Document?.Body?.ChildElements.Where(x =>
                _etalonRequestPattern.IsMatch(x.InnerText)
                || _etalonResponsePattern.IsMatch(x.InnerText)
                || _testPattern.IsMatch(x.InnerText)
            )
            .ToArray();

        _entry.TotalCount = templateParagraphs?.Length ?? 0;

        if (_entry.TotalCount == 0)
        {
            throw new Exception(
                "Нет шаблонов для заполнения. Проверьте исходный файл на существование шаблонов типа #a/1/method или #b/1"
            );
        }

        InitStyles(newDoc);
        ProcessTemplates(templateParagraphs);
    }

    private void ProcessTemplates(OpenXmlElement[] templateParagraphs)
    {
        var currentStep = 0;

        foreach (var openXmlElement in templateParagraphs)
        {
            if (openXmlElement is not Paragraph p) continue;

            ProcessTemplate(p);

            _worker.ReportProgress(++currentStep);
        }
    }

    private ProcessType? ProcessMatch(string text)
    {
        if (_etalonRequestPattern.IsMatch(text))
        {
            return ProcessType.EtalonRequest;
        }

        if (_etalonResponsePattern.IsMatch(text))
        {
            return ProcessType.EtalonResponse;
        }

        if (_testPattern.IsMatch(text))
        {
            return ProcessType.Test;
        }

        return null;
    }

    private XDocument GetXmlElement(string xmlPath)
    {
        var xmlDocument = new XmlDocument();
        xmlDocument.Load(xmlPath);
        var element = XDocument.Parse(xmlDocument.OuterXml);

        return element;
    }

    private string GetXmlPath(string xmlNumber, ProcessType processType)
    {
        if (string.IsNullOrEmpty(xmlNumber))
        {
            return string.Empty;
        }

        return processType switch
        {
            ProcessType.EtalonRequest => $"{_entry.EtalonFolder}/{xmlNumber}/Request.xml",
            ProcessType.EtalonResponse => $"{_entry.EtalonFolder}/{xmlNumber}/Response.xml",
            ProcessType.Test => $"{_entry.TestFolder}/{xmlNumber}.xsl",
            _ => string.Empty,
        };
    }

    private void ProcessTemplate(Paragraph p)
    {
        var text = p.InnerText;
        _entry.CurrentTemplate = text;

        p.RemoveAllChildren<Run>();

        var processType = ProcessMatch(text) ?? throw new Exception("Не найден вид шаблона");
        var xmlPath = GetXmlPath(text.Split("/")[1], processType);
        var element = GetXmlElement(xmlPath);
        SetXml(p, element.Elements(), 0);
    }

    private void InitStyles(WordprocessingDocument newDoc)
    {
        var part = newDoc?.MainDocumentPart?.StyleDefinitionsPart;

        if (part == null)
        {
            part = newDoc?.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
            var root = new Styles();
            root.Save(part);
        }

        var styles = part.Styles;

        var blueTagStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "blueTag",
            CustomStyle = true,
        };

        var plainTextStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "plainText",
            CustomStyle = true,
        };

        var styleName1 = new StyleName
        {
            Val = "Blue tag",
        };

        var linkedStyle1 = new LinkedStyle
        {
            Val = "linkedBlue",
        };

        blueTagStyle.Append(styleName1);
        blueTagStyle.Append(linkedStyle1);

        var runStyle = new StyleRunProperties();

        var color = new Color
        {
            ThemeColor = ThemeColorValues.Accent1,
        };

        var font = new RunFonts
        {
            Ascii = "Times New Roman",
        };

        var fontSize = new FontSize
        {
            Val = "24",
        };

        runStyle.Append(color);
        runStyle.Append(font);
        runStyle.Append(fontSize);
        blueTagStyle.Append(runStyle);
        styles?.Append(blueTagStyle);

        var styleName2 = new StyleName
        {
            Val = "Plain text",
        };

        var linkedStyle2 = new LinkedStyle
        {
            Val = "linkedPlain",
        };

        var color2 = new Color
        {
            ThemeColor = ThemeColorValues.Text1,
        };

        var font2 = new RunFonts
        {
            Ascii = "Times New Roman",
        };

        var fontSize2 = new FontSize
        {
            Val = "24",
        };

        plainTextStyle.Append(styleName2);
        plainTextStyle.Append(linkedStyle2);
        var runStyle2 = new StyleRunProperties();
        runStyle2.Append(color2);
        runStyle2.Append(font2);
        runStyle2.Append(fontSize2);
        plainTextStyle.Append(runStyle2);
        styles?.Append(plainTextStyle);
    }

    private void SetXml(Paragraph p, IEnumerable<XElement> elements, int level)
    {
        foreach (var elem in elements)
        {
            var indent = level * 2;
            var newtext = new Text();
            var newrun = new Run();
            var runProp = newrun.RunProperties ?? (newrun.RunProperties = new RunProperties());

            runProp.RunStyle = new RunStyle
            {
                Val = "blueTag",
            };

            newtext.Space = SpaceProcessingModeValues.Preserve;

            newtext.Text =
                $"{new string(' ', indent)}<{elem.Name.LocalName}{(elem.Attributes().Any() ? " " : "")}{string.Join(" ", elem.Attributes())}>";

            if(level > 0)
                newrun.Append(new Break());

            newrun.Append(newtext);
            p.Append(newrun);

            SetXml(p, elem.Elements(), level + 1);

            if (!elem.Elements()
                    .Any()
                && elem.Value != null)
            {
                newtext = new Text();
                newrun = new Run();
                runProp = newrun.RunProperties ?? (newrun.RunProperties = new RunProperties());

                runProp.RunStyle = new RunStyle
                {
                    Val = "plainText",
                };

                newtext.Space = SpaceProcessingModeValues.Preserve;
                newtext.Text = elem.Value;
                newrun.Append(newtext);
                p.Append(newrun);
                indent = 0;
            }

            newtext = new Text();
            newrun = new Run();
            runProp = newrun.RunProperties ?? (newrun.RunProperties = new RunProperties());

            newtext.Space = SpaceProcessingModeValues.Preserve;

            runProp.RunStyle = new RunStyle
            {
                Val = "blueTag",
            };

            newtext.Text = $"{new string(' ', indent)}</{elem.Name.LocalName}>";

            if (elem.Elements()
                .Any())
            {
                newrun.Append(new Break());
            }

            newrun.Append(newtext);
            p.Append(newrun);
        }
    }

    #region Worker

    private void InitWorker()
    {
        _worker = new BackgroundWorker
        {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true,
        };

        _worker.DoWork += worker_DoWork;
        _worker.ProgressChanged += worker_ProgressChanged;
        _worker.RunWorkerCompleted += _worker_RunWorkerCompleted;
    }

    private void worker_DoWork(object sender, DoWorkEventArgs e)
    {
        if (_entry == null)
        {
            return;
        }

        _entry.StartVisible = Visibility.Collapsed;
        ExecuteInternal(e);
    }

    private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
        if (_entry != null)
        {
            _entry.Progress = e.ProgressPercentage;
        }
    }

    public void worker_Stop()
    {
        if (_worker is {WorkerSupportsCancellation: true, CancellationPending: false,})
        {
            _worker.CancelAsync();
        }
    }

    private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        if (_entry != null)
        {
            _entry.StartVisible = Visibility.Visible;
            _entry.Progress = 0;
            _entry.TotalCount = 0;
            _entry.CurrentTemplate = string.Empty;
        }

        if (e.Error != null)
        {
            MessageBox.Show(e.Error.Message, "Ошибка");

            return;
        }

        MessageBox.Show($"Задача {(!e.Cancelled ? "выполнена" : "отменена")}", "Завершение");
    }

    #endregion

    private enum ProcessType
    {
        EtalonRequest,
        EtalonResponse,
        Test,
    }
}