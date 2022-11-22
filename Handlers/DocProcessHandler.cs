using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using tff.main.Extensions;
using tff.main.Models;
using MessageBox = System.Windows.Forms.MessageBox;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace tff.main.Handlers;

public class DocProcessHandler
{
    private readonly Entry _entry;
    private readonly Regex _etalonRequestPattern;
    private readonly Regex _etalonResponsePattern;
    private readonly string _minfinUrn;
    private readonly Regex _testPattern;
    private readonly Regex _xsdPattern;
    private BackgroundWorker _worker;

    public DocProcessHandler(Entry entry)
    {
        InitWorker();
        _etalonRequestPattern = new Regex("#a/\\d{1,}/request");
        _etalonResponsePattern = new Regex("#a/\\d{1,}/response");
        _testPattern = new Regex("#b/\\d{1,}");
        _xsdPattern = new Regex("#xsd");
        _minfinUrn = "urn://x-artefacts-minfin-unp/1.11.0";
        _entry = entry;
    }

    public void Execute()
    {
        _worker.RunWorkerAsync();
    }

    private void ExecuteInternal(WordprocessingDocument newDoc)
    {
        var templateParagraphsQuery = newDoc.MainDocumentPart?.Document.Body?.ChildElements.Where(x =>
                _etalonRequestPattern.IsMatch(x.InnerText)
             || _etalonResponsePattern.IsMatch(x.InnerText)
             || _testPattern.IsMatch(x.InnerText)
            )
         ?? throw new InvalidOperationException(
                "Нет шаблонов для заполнения. Проверьте исходный файл на существование шаблонов типа #a/1/method или #b/1"
            );

        if (_entry.EtalonFolder == null)
        {
            templateParagraphsQuery = templateParagraphsQuery.Where(x => _testPattern.IsMatch(x.InnerText));
        }
        else if (_entry.TestFolder == null)
        {
            templateParagraphsQuery = templateParagraphsQuery.Where(x => !_testPattern.IsMatch(x.InnerText));
        }

        var templateParagraphs = templateParagraphsQuery.ToArray();

        ProcessTemplates(templateParagraphs);
    }

    private int GetProgress(int currentStep)
    {
        return (int) decimal.Round(currentStep / (decimal) _entry.TotalCount * 100);
    }

    private enum ProcessType
    {
        EtalonRequest,
        EtalonResponse,
        Test,
    }

    #region XSD

    private void ExecuteInternalXsd(WordprocessingDocument newDoc)
    {
        using var fs = new FileStream(_entry.TargetXsdFile, FileMode.Open);

        var reader = XmlReader.Create(fs,
            new XmlReaderSettings
            {
                IgnoreWhitespace = true,
                IgnoreProcessingInstructions = true,
            }
        );

        var schema = XmlSchema.Read(reader, ValidationCallback)
         ?? throw new InvalidOperationException("Невозможно прочитать файл XSD. Проверьте его корректность.");

        _entry.CurrentTemplate = "Подготовка данных...";

        var (xsdDescriptionsElements, xsdDescriptionsComplexes) = GetXsdDescriptions(schema);

        OpenXmlElement nextElement;

        if (xsdDescriptionsElements.Length > 0 && xsdDescriptionsComplexes.Length > 0)
        {
            _entry.TotalCount = xsdDescriptionsElements.Length + xsdDescriptionsComplexes.Length;
            _entry.Progress = 0;

            nextElement = newDoc.MainDocumentPart?.Document.Body?.ChildElements.FirstOrDefault(x =>
                _xsdPattern.IsMatch(x.InnerText)
            );

            if (nextElement == null)
            {
                throw new InvalidOperationException($"Не найден шаблон {_xsdPattern} для XSD документа");
            }

            nextElement.GetFirstChild<Run>().GetFirstChild<Text>().Text = "TUTA ZAMENA";
        }
        else
        {
            throw new InvalidOperationException("Не найдены элементы в файле XSD. Проверьте его корректность.");
        }

        var counter = 0;
        var allElements = xsdDescriptionsElements.Concat(xsdDescriptionsComplexes).ToArray();

        XsdExtension.SetReferenceComment(allElements);

        var paragraphsWithHyperlink = new List<KeyValuePair<Paragraph, XsdDescription>>();

        //элементы
        WalkNestedElements(ref nextElement,
            allElements,
            paragraphsWithHyperlink,
            xsdDescriptionsElements.Where(x => x.Parent == null).ToArray(),
            1,
            2,
            5,
            ref counter
        );

        nextElement = InsertTitle(nextElement,
            "Описание комплексных типов (при наличии)",
            "4.3",
            1,
            2,
            5
        );

        //комплексные типы
        WalkNestedElements(ref nextElement,
            allElements,
            paragraphsWithHyperlink,
            xsdDescriptionsComplexes.Where(x => x.Parent == null).ToArray(),
            0,
            30,
            10,
            ref counter
        );

        foreach (var (child, parent) in paragraphsWithHyperlink)
        {
            var run = child.GetFirstChild<Run>();

            if (run == null)
            {
                continue;
            }

            XsdExtension.AddHyperlink(run, parent.Number);
        }

        _entry.CurrentTemplate = "Сохранение файла...";
    }

    private Paragraph InsertTitle(OpenXmlElement nextElement,
        string title,
        string number,
        int level,
        int styleId,
        int numId)
    {
        var text = new Text();
        var run = new Run();
        var paragraph = new Paragraph();

        var parStyle = new ParagraphStyleId
        {
            Val = $"{styleId}",
        };

        paragraph.ParagraphProperties = new ParagraphProperties
        {
            ParagraphStyleId = parStyle,
            NumberingProperties = new NumberingProperties
            {
                NumberingLevelReference = new NumberingLevelReference
                {
                    Val = level,
                },
                NumberingId = new NumberingId
                {
                    Val = numId,
                },
            },
        };

        text.Text = title;
        run.AppendChild(text);
        paragraph.AppendChild(run);

        nextElement.InsertAfterSelf(paragraph);

        paragraph.AddBookmark(number);

        return paragraph;
    }

    private void WalkNestedElements(ref OpenXmlElement nextElement,
        XsdDescription[] allElements,
        List<KeyValuePair<Paragraph, XsdDescription>> paragraphsWithHyperlink,
        XsdDescription[] elements,
        int level,
        int styleId,
        int numId,
        ref int progress)
    {
        foreach (var xsdDescription in elements)
        {
            ++progress;

            _entry.CurrentTemplate = !string.IsNullOrEmpty(xsdDescription.Field)
                ? $"{xsdDescription.Field}"
                : _entry.CurrentTemplate;

            _worker.ReportProgress(GetProgress(progress));

            var childs = allElements.Where(x => x.Parent == xsdDescription);
            var xsdDescriptions = childs as XsdDescription[] ?? childs.ToArray();
            childs = Array.Empty<object>() as IEnumerable<XsdDescription>;

            if (!xsdDescriptions.Any() || nextElement == null)
            {
                continue;
            }

            var paragraph = InsertTitle(nextElement,
                xsdDescription.Title,
                xsdDescription.Number,
                level,
                styleId,
                numId
            );

            var tableProperties = new TableProperties(new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = "auto",
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = "auto",
                        Space = 0,
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = "auto",
                        Space = 0,
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = "auto",
                        Space = 0,
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = "auto",
                        Space = 0,
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = "auto",
                        Space = 0,
                    },
                    new TableWidth
                    {
                        Type = TableWidthUnitValues.Pct,
                        Width = "5060",
                    },
                    new TableLayout
                    {
                        Type = TableLayoutValues.Fixed,
                    },
                    new TableLook
                    {
                        FirstRow = true,
                        LastRow = false,
                        FirstColumn = true,
                        LastColumn = false,
                        NoHorizontalBand = false,
                        NoVerticalBand = true,
                    },
                    new TableStyle
                    {
                        Val = "TableNormal",
                    }
                )
            );

            var tableGrid = new TableGrid(new GridColumn
                {
                    Width = "1000",
                },
                new GridColumn
                {
                    Width = "2000",
                },
                new GridColumn
                {
                    Width = "2000",
                },
                new GridColumn
                {
                    Width = "1500",
                },
                new GridColumn
                {
                    Width = "1500",
                },
                new GridColumn
                {
                    Width = "2000",
                }
            );

            var table = new Table();

            table.AppendChild(tableProperties);
            table.AppendChild(tableGrid);

            paragraph.InsertAfterSelf(table);
            table.InsertTableHeader();

            paragraph.AddBookmark(xsdDescription.Number);

            var counter = 0;

            foreach (var child in xsdDescriptions)
                table.InsertTableRow(paragraphsWithHyperlink, allElements, child, ++counter);

            nextElement = paragraph.NextSibling();

            WalkNestedElements(ref nextElement,
                allElements,
                paragraphsWithHyperlink,
                xsdDescriptions.ToArray(),
                level,
                styleId,
                numId,
                ref progress
            );
        }
    }

    private static void ValidationCallback(object sender, ValidationEventArgs args) { }

    private (XsdDescription[], XsdDescription[]) GetXsdDescriptions(XmlSchema schema)
    {
        var counter = 0;
        var complexCounter = 0;
        var elements = new List<XsdDescription>();
        var complexes = new List<XsdDescription>();

        foreach (var item in schema.Items)
            switch (item)
            {
                case XmlSchemaElement element:
                {
                    ++counter;
                    var name = element.Name ?? string.Empty;
                    var type = element.SchemaTypeName.IsEmpty ? string.Empty : element.SchemaTypeName.Name;
                    var annotation = GetAnnotation(element);

                    var title = "Описание полей ";

                    switch (name)
                    {
                        case "UNPDocument":
                            title += "запроса";

                            break;
                        case "UNPDocumentResponse":
                            title += "ответа на запрос";

                            break;
                    }

                    var xsdEntity = new XsdDescription
                    {
                        Number = $"4.{counter}",
                        Title = title,
                        Annotation = annotation,
                        IsRequired = true,
                        Field = name,
                        Parent = null,
                        TypeOrFillMethod = type,
                    };

                    elements.Add(xsdEntity);

                    WalkChildNodes(xsdEntity, element, elements, ref counter);

                    break;
                }
                case XmlSchemaComplexType complexType:
                {
                    ++complexCounter;

                    var annotation = GetAnnotation(complexType);
                    var name = complexType.Name ?? string.Empty;
                    var number = $"4.{counter + 1}.{complexCounter}";
                    var title = $"Описание комплексного элемента {name} ({annotation})";

                    var xsdEntity = new XsdDescription
                    {
                        Number = number,
                        Title = title,
                        Annotation = annotation,
                        IsRequired = true,
                        Field = name,
                        Parent = null,
                        TypeOrFillMethod = string.Empty,
                    };

                    complexes.Add(xsdEntity);

                    WalkChildNodes(xsdEntity, complexType.Particle, complexes, ref complexCounter);

                    break;
                }
            }

        return (elements.ToArray(), complexes.ToArray());
    }

    private void WalkChildNodes(XsdDescription parent, XmlSchemaParticle particle, IList complexes, ref int counter)
    {
        if (particle is XmlSchemaElement xmlSchemaElement)
        {
            particle = new XmlSchemaComplexType().Particle = new XmlSchemaSequence
            {
                Items =
                {
                    new XmlSchemaElement
                    {
                        SchemaTypeName = xmlSchemaElement.SchemaTypeName,
                        Name = xmlSchemaElement.Name,
                        Parent = xmlSchemaElement.Parent,
                        SchemaType = xmlSchemaElement.SchemaType,
                        Annotation = xmlSchemaElement.Annotation,
                        MinOccurs = xmlSchemaElement.MinOccurs,
                    },
                },
            };
        }

        if (particle is not XmlSchemaGroupBase schemaGroupBase)
        {
            return;
        }

        var nodes = schemaGroupBase.Items;

        foreach (var node in nodes)
            switch (node)
            {
                case XmlSchemaElement element:
                {
                    var name = element.Name ?? string.Empty;

                    var annotation = GetAnnotation(element);

                    var type = element.SchemaTypeName.IsEmpty
                        ? string.Empty
                        : $"{(element.SchemaTypeName.Namespace == _minfinUrn ? string.Empty : "xs:")}{element.SchemaTypeName.Name}";

                    var minOccurs = element.MinOccurs;
                    var comment = string.Empty;

                    if (annotation.Contains('\n'))
                    {
                        var subStringLength = annotation.IndexOf('\n');

                        comment = annotation.Remove(0, subStringLength).Trim(' ', '\n');

                        annotation = annotation[..subStringLength].Trim('\n');
                    }

                    var xsdEntity = new XsdDescription
                    {
                        Annotation = annotation,
                        IsRequired = particle is not XmlSchemaChoice && minOccurs != 0,
                        Field = name,
                        Parent = parent,
                        TypeOrFillMethod = type,
                        Comment = comment,
                    };

                    if (element.SchemaType is XmlSchemaComplexType {Particle: { },} && !string.IsNullOrEmpty(annotation))
                    {
                        complexes.Add(xsdEntity);
                    }
                    else
                    {
                        complexes.Add(xsdEntity);
                    }

                    switch (element.SchemaType)
                    {
                        case XmlSchemaComplexType {Particle: { },} complexType:
                        {
                            counter++;

                            var isComplex = string.IsNullOrEmpty(annotation);

                            if (isComplex)
                            {
                                annotation = GetAnnotation(complexType);
                            }

                            xsdEntity.Title =
                                $"Описание {(isComplex ? "комплексного" : "составного")} элемента {name}{(string.IsNullOrEmpty(annotation) ? "" : $" ({annotation.Trim(':')})")}";

                            xsdEntity.TypeOrFillMethod = string.IsNullOrEmpty(type)
                                ? isComplex ? "Комплексный элемент" : "Составной элемент"
                                : type;

                            if (!isComplex)
                            {
                                annotation = annotation + '\n' + comment;
                            }

                            xsdEntity.Number = $"4.3.{counter}";
                            xsdEntity.Annotation = annotation;

                            WalkChildNodes(string.IsNullOrEmpty(annotation) ? parent : xsdEntity,
                                complexType.Particle,
                                complexes,
                                ref counter
                            );

                            break;
                        }

                        case XmlSchemaSimpleType simpleType:
                        {
                            var simpleElementType = string.Empty;
                            var simpleComment = string.Empty;

                            if (simpleType.Content is XmlSchemaSimpleTypeRestriction simpleTypeContent)
                            {
                                simpleElementType = $"xs:{simpleTypeContent.BaseTypeName.Name}";

                                if (simpleTypeContent.Facets.Count > 0)
                                {
                                    var maxLength = "";
                                    var minLength = "";

                                    foreach (XmlSchemaFacet facet in simpleTypeContent.Facets)
                                    {
                                        var value = facet.Value;

                                        switch (facet.GetType().Name)
                                        {
                                            case nameof(XmlSchemaMaxLengthFacet):
                                                maxLength = value;

                                                break;

                                            case nameof(XmlSchemaMinLengthFacet):
                                                minLength = value;

                                                break;
                                        }
                                    }

                                    simpleComment = string.IsNullOrEmpty(minLength) || string.IsNullOrEmpty(maxLength)
                                        ? $"Длина:{(!string.IsNullOrEmpty(minLength) ? " от " + minLength : "")}{(!string.IsNullOrEmpty(maxLength) ? " до " + maxLength : "")} знаков"
                                        : "";
                                }
                            }

                            xsdEntity.Comment += $"\n{simpleComment}";
                            xsdEntity.Comment = xsdEntity.Comment.Trim('\n', '(', ')');
                            xsdEntity.TypeOrFillMethod = simpleElementType;

                            break;
                        }
                    }

                    break;
                }
                case XmlSchemaChoice choice:
                    WalkChildNodes(parent, choice, complexes, ref counter);

                    break;
            }
    }

    private string GetAnnotation(XmlSchemaAnnotated schemaObject)
    {
        return (schemaObject.Annotation?.Items[0] as XmlSchemaDocumentation)?.Markup?[0]?.Value?.TrimEnd(':')
         ?? string.Empty;
    }

    #endregion

    #region Word

    private void ProcessTemplates(OpenXmlElement[] templateParagraphs)
    {
        var currentStep = 0;
        _entry.TotalCount = templateParagraphs.Length;

        foreach (var openXmlElement in templateParagraphs)
        {
            if (openXmlElement is not Paragraph p)
            {
                continue;
            }

            ProcessTemplate(p);

            _worker.ReportProgress(GetProgress(++currentStep));
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

            if (part == null)
            {
                return;
            }

            var root = new Styles();
            root.Save(part);
        }

        var styles = part.Styles;

        #region bluetag

        var blueTagStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "blueTag",
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

        #endregion

        #region plainText

        var plainTextStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "plainText",
            CustomStyle = true,
        };

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

        #endregion

        styles?.Append(plainTextStyle);
        styles?.Append(blueTagStyle);
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

            if (level > 0)
            {
                newrun.Append(new Break());
            }

            newrun.Append(newtext);
            p.Append(newrun);

            SetXml(p, elem.Elements(), level + 1);

            if (!elem.Elements().Any() && elem.Value != null)
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

            if (elem.Elements().Any())
            {
                newrun.Append(new Break());
            }

            newrun.Append(newtext);
            p.Append(newrun);
        }
    }

    #endregion

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

        using var newDoc = GetNewWordDocument();

        InitStyles(newDoc);

        if (!string.IsNullOrEmpty(_entry.EtalonFolder) || !string.IsNullOrEmpty(_entry.TestFolder))
        {
            ExecuteInternal(newDoc);
        }

        if (!string.IsNullOrEmpty(_entry.TargetXsdFile))
        {
            ExecuteInternalXsd(newDoc);
        }

        if (!newDoc.AutoSave)
        {
            newDoc.Save();
        }
    }

    private WordprocessingDocument GetNewWordDocument()
    {
        var fileInfo = new FileInfo(_entry.TargetFile ?? throw new InvalidOperationException());
        var newFilePath = $"{_entry.SavePath}\\{fileInfo.Name.Split(".")[0]}_ГОТОВЫЙ{fileInfo.Extension}";

        var oldDoc = WordprocessingDocument.Open(_entry.TargetFile,
            false,
            new OpenSettings
            {
                AutoSave = false,
                MarkupCompatibilityProcessSettings =
                    new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts,
                        FileFormatVersions.Office2010
                    ),
            }
        );

        using (oldDoc.Clone(newFilePath,
                false,
                new OpenSettings
                {
                    MarkupCompatibilityProcessSettings =
                        new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts,
                            FileFormatVersions.Office2010
                        ),
                }
            ))
        {
            oldDoc.Dispose();
        }

        return WordprocessingDocument.Open(newFilePath,
            true,
            new OpenSettings
            {
                MarkupCompatibilityProcessSettings =
                    new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts,
                        FileFormatVersions.Office2010
                    ),
            }
        );
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
}