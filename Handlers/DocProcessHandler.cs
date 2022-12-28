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
using DocumentFormat.OpenXml.Validation;
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
    private string _minfinUrn;
    private readonly Regex _testPattern;
    private readonly Regex _xsdPattern;
    private readonly Regex _xsdListPattern;
    private readonly Regex _urnPathPattern;
    private readonly Regex _numTestPattern;
    private readonly Regex _versionPattern;
    private readonly Regex _xsdVersionPattern;
    private BackgroundWorker _worker;

    public DocProcessHandler(Entry entry)
    {
        InitWorker();
        _etalonRequestPattern = new Regex("#a/\\d{1,}/request");
        _etalonResponsePattern = new Regex("#a/\\d{1,}/response");
        _testPattern = new Regex("#b/\\d{1,}");
        _xsdPattern = new Regex("#xsd");
        _xsdListPattern = new Regex("#xmlList");
        _urnPathPattern = new Regex("#urn");
        _numTestPattern = new Regex("#num_test");
        _versionPattern = new Regex("#version");
        _xsdVersionPattern = new Regex("#xsdVersion");
        _entry = entry;
    }

    public void Execute()
    {
        _worker.RunWorkerAsync();
    }

    private void ExecuteInternal(WordprocessingDocument newDoc)
    {
        var templateParagraphsQuery = newDoc.MainDocumentPart?.Document.Body?.ChildElements.Where(x =>
                                          PatternMatch(_etalonRequestPattern, x.InnerText)
                                          || PatternMatch(_etalonResponsePattern, x.InnerText)
                                          || PatternMatch(_testPattern, x.InnerText)
                                      )
                                      ?? throw new InvalidOperationException(
                                          "Нет шаблонов для заполнения. Проверьте исходный файл на существование шаблонов типа #a/1/method или #b/1"
                                      );

        if (_entry.EtalonFolder == null)
        {
            templateParagraphsQuery = templateParagraphsQuery.Where(x => PatternMatch(_testPattern, x.InnerText));
        }
        else if (_entry.TestFolder == null)
        {
            templateParagraphsQuery = templateParagraphsQuery.Where(x => !PatternMatch(_testPattern, x.InnerText));
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

    private void ExecuteInternalXsd(WordprocessingDocument newDoc, int numId)
    {
        using var fs = new FileStream(_entry.TargetXsdFile, FileMode.Open);

        using (var reader = XmlReader.Create(fs,
                   new XmlReaderSettings
                   {
                       IgnoreWhitespace = true,
                       IgnoreProcessingInstructions = true,
                   }
               ))
        {
            var schema = XmlSchema.Read(reader, ValidationCallback)
                         ?? throw new InvalidOperationException(
                             "Невозможно прочитать файл XSD. Проверьте его корректность.");

            _entry.CurrentTemplate = "Подготовка данных...";
            _minfinUrn = schema.TargetNamespace;

            var (xsdDescriptionsElements, xsdDescriptionsComplexes) = GetXsdDescriptions(schema);

            OpenXmlElement nextElement;

            if (xsdDescriptionsElements.Length > 0 && xsdDescriptionsComplexes.Length > 0)
            {
                _entry.TotalCount = xsdDescriptionsElements.Length + xsdDescriptionsComplexes.Length;
                _entry.Progress = 0;

                nextElement = newDoc.MainDocumentPart?.Document.Body?.ChildElements.FirstOrDefault(x =>
                    _xsdPattern.IsMatch(x.InnerText) && !_xsdVersionPattern.IsMatch(x.InnerText)
                );

                if (nextElement == null)
                {
                    throw new InvalidOperationException($"Не найден шаблон {_xsdPattern} для XSD документа");
                }

                nextElement.GetFirstChild<Run>().Remove();
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
                "2",
                5,
                ref counter
            );

            nextElement = XsdExtension.InsertTitle(nextElement,
                "Описание комплексных типов полей (при наличии)",
                "4.3",
                1,
                "2",
                5
            );

            //комплексные типы
            WalkNestedElements(ref nextElement,
                allElements,
                paragraphsWithHyperlink,
                xsdDescriptionsComplexes.Where(x => x.Parent == null).ToArray(),
                2,
                "Numbering_4_3_x",
                numId,
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
        }

        var xsdListing =
            newDoc.MainDocumentPart?.Document.Body?.ChildElements.FirstOrDefault(x =>
                PatternMatch(_xsdListPattern, x.InnerText));
        fs.Seek(0, SeekOrigin.Begin);
        var xmlDocument = new XmlDocument();
        xmlDocument.Load(fs);
        var element = XDocument.Parse(xmlDocument.OuterXml);

        InsertXsdInnerText(xsdListing, element);

        var namespaceUrnArray = newDoc.MainDocumentPart?.Document.Body?.ChildElements
            .Where(x => PatternMatch(_urnPathPattern, x.InnerText))
            .ToArray();

        SetNamespaceUrn(namespaceUrnArray);

        SetVersion(newDoc.MainDocumentPart?.Document.Body?.ChildElements.Where(x =>
                PatternMatch(_versionPattern, x.InnerText)
            ).ToArray(),
            GetVersion()
        );

        SetXsdVersion(newDoc.MainDocumentPart?.Document.Body?.ChildElements.FirstOrDefault(x =>
                PatternMatch(_xsdVersionPattern, x.InnerText)
            ),
            GetVersion());

        _entry.CurrentTemplate = "Сохранение файла...";
    }

    private string GetVersion() => _minfinUrn.Split('/')[^1];

    private void SetXsdVersion(OpenXmlElement element, string version)
    {
        if (element is not Paragraph p || string.IsNullOrWhiteSpace(version))
        {
            return;
        }

        var neighbourElement = p.ChildElements.FirstOrDefault(x => PatternMatch(_xsdVersionPattern, x.InnerText));

        var run = new Run(new Text(version));
        neighbourElement?.InsertAfterSelf(run);
        neighbourElement?.Remove();
    }

    private void SetVersion(OpenXmlElement[] elements, string version)
    {
        foreach (var element in elements)
        {
            if (element is not Paragraph p || string.IsNullOrWhiteSpace(version))
            {
                return;
            }

            p.RemoveAllChildren<Run>();

            var run = new Run(new Text($"Версия: {version}"));

            run.RunProperties ??= new RunProperties
            {
                Bold = new Bold {Val = OnOffValue.FromBoolean(true)},
                FontSize = new FontSize
                {
                    Val = "24",
                },
                RunFonts = new RunFonts
                {
                    Ascii = "Times New Roman",
                    ComplexScript = "Times New Roman",
                    EastAsia = "Times New Roman",
                    HighAnsi = "Times New Roman",
                },
            };

            p.Append(run,
                new Run(new Break
                    {
                        Type = new EnumValue<BreakValues>(BreakValues.Page),
                    }
                )
            );
        }
    }

    private void SetNamespaceUrn(OpenXmlElement[] elements)
    {
        _entry.CurrentTemplate = _urnPathPattern.ToString();

        foreach (var openXmlElement in elements)
        {
            if (openXmlElement is Table t)
            {
                var p = t.ChildElements.FirstOrDefault(x => _urnPathPattern.IsMatch(x.InnerText))
                    ?.ChildElements.FirstOrDefault(x => _urnPathPattern.IsMatch(x.InnerText))?
                    .GetFirstChild<Paragraph>();

                p?.RemoveAllChildren<Run>();

                var run = new Run();
                var text = new Text();

                run.RunProperties = new RunProperties
                {
                    RunStyle = new RunStyle
                    {
                        Val = "plainText",
                    },
                };

                text.Text = _minfinUrn;
                run.Append(text);
                p?.Append(run);
            }
        }
    }

    /// <summary>
    /// рекурсивный обход вложенных элементов
    /// </summary>
    /// <param name="nextElement">следующий элемент</param>
    /// <param name="allElements">все элементы</param>
    /// <param name="paragraphsWithHyperlink">список параграфов с ссылками</param>
    /// <param name="elements">список элементов</param>
    /// <param name="level">уровень вложенности</param>
    /// <param name="styleId">Id задаваемого стиля</param>
    /// <param name="numId">Id стиля нумерации</param>
    /// <param name="progress">прогресс выполнения</param>
    private void WalkNestedElements(ref OpenXmlElement nextElement,
        XsdDescription[] allElements,
        List<KeyValuePair<Paragraph, XsdDescription>> paragraphsWithHyperlink,
        XsdDescription[] elements,
        int level,
        string styleId,
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

            if (!xsdDescriptions.Any() || nextElement == null)
            {
                continue;
            }

            var paragraph = XsdExtension.InsertTitle(nextElement,
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

    private void InsertXsdInnerText(OpenXmlElement openXmlElement, XDocument element)
    {
        if (openXmlElement is not Table t)
        {
            return;
        }

        var p = t.GetFirstChild<TableRow>().GetFirstChild<TableCell>().GetFirstChild<Paragraph>();
        var text = t.InnerText;
        _entry.CurrentTemplate = text;
        p.RemoveAllChildren<Run>();

        SetXml(p, element.Elements(), 0, false);
    }

    private static void ValidationCallback(object sender, ValidationEventArgs args)
    { }

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
                    var annotation = XsdExtension.GetAnnotation(element);

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

                    var annotation = XsdExtension.GetAnnotation(complexType);
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

                    var annotation = XsdExtension.GetAnnotation(element);

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
                            var isComplex = string.IsNullOrEmpty(annotation);

                            if (isComplex)
                            {
                                annotation = XsdExtension.GetAnnotation(complexType);
                            }

                            if (string.IsNullOrEmpty(annotation))
                            {
                                complexes.Remove(xsdEntity);
                                WalkChildNodes(parent, complexType.Particle, complexes, ref counter);

                                continue;
                            }

                            counter++;

                            xsdEntity.Title =
                                $"Описание {(isComplex ? "комплексного" : "составного")} типа {name}{(string.IsNullOrEmpty(annotation) ? "" : $" ({annotation.Trim(':')})")}";

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

    #endregion

    #region Word

    private void ProcessTemplates(OpenXmlElement[] templateParagraphs)
    {
        var currentStep = 0;
        _entry.TotalCount = templateParagraphs.Length;

        foreach (var openXmlElement in templateParagraphs)
        {
            if (openXmlElement is not Table t)
            {
                continue;
            }

            if (_worker.CancellationPending)
            {
                break;
            }

            var p = t.GetFirstChild<TableRow>().GetFirstChild<TableCell>().GetFirstChild<Paragraph>();

            ProcessTemplate(p);

            _worker.ReportProgress(GetProgress(++currentStep));
        }
    }

    private ProcessType? ProcessMatch(string text)
    {
        if (PatternMatch(_etalonRequestPattern, text))
        {
            return ProcessType.EtalonRequest;
        }

        if (PatternMatch(_etalonResponsePattern, text))
        {
            return ProcessType.EtalonResponse;
        }

        if (PatternMatch(_testPattern, text))
        {
            return ProcessType.Test;
        }

        return null;
    }

    private bool PatternMatch(Regex regex, string text) => regex.IsMatch(text);

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
            ProcessType.EtalonRequest  => $"{_entry.EtalonFolder}/{xmlNumber}/Request.xml",
            ProcessType.EtalonResponse => $"{_entry.EtalonFolder}/{xmlNumber}/Response.xml",
            ProcessType.Test           => $"{_entry.TestFolder}/{xmlNumber}.xsl",
            _                          => string.Empty,
        };
    }

    private void ProcessTemplate(Paragraph p)
    {
        var text = p.InnerText;
        var xmlNumber = text.Split("/")[1];
        _entry.CurrentTemplate = text;

        p.RemoveAllChildren<Run>();

        var processType = ProcessMatch(text) ?? throw new Exception("Не найден вид шаблона");
        var xmlPath = GetXmlPath(xmlNumber, processType);
        var element = GetXmlElement(xmlPath);
        SetXml(p, element.Elements(), 0);

        var testTable = p.Parent.Parent.Parent.NextSibling<Table>();

        if (processType != ProcessType.Test)
        {
            return;
        }

        var paragraph = testTable.ChildElements.FirstOrDefault(x => _numTestPattern.IsMatch(x.InnerText))
            ?.ChildElements.FirstOrDefault(x => _numTestPattern.IsMatch(x.InnerText))
            ?.GetFirstChild<Paragraph>();

        paragraph?.RemoveAllChildren<Run>();

        var run = new Run(new Text($"ответ - {xmlNumber}.xls"));

        run.RunProperties ??= new RunProperties
        {
            RunStyle = new RunStyle
            {
                Val = "plainText",
            },
        };

        paragraph?.Append(run);
    }

    private (int, int) InitStyles(WordprocessingDocument newDoc)
    {
        var part = newDoc?.MainDocumentPart?.StyleDefinitionsPart;
        var numPart = newDoc?.MainDocumentPart?.NumberingDefinitionsPart;

        if (part == null)
        {
            part = newDoc?.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();

            if (part == null)
            {
                return (-1, -1);
            }

            var root = new Styles();
            root.Save(part);
        }

        if (numPart == null)
        {
            numPart = newDoc.MainDocumentPart?.AddNewPart<NumberingDefinitionsPart>();

            if (numPart == null)
            {
                return (-1, -1);
            }

            var root = new Numbering();
            root.Save(numPart);
        }

        var styles = part.Styles;
        var numStyles = numPart.Numbering;

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

        blueTagStyle.AppendChild(styleName1);
        blueTagStyle.AppendChild(linkedStyle1);
        var runStyle = new StyleRunProperties();

        var color = new Color
        {
            ThemeColor = ThemeColorValues.Accent1,
        };

        var font = new RunFonts
        {
            Ascii = "Times New Roman",
            EastAsia = "Times New Roman",
            ComplexScript = "Times New Roman",
            HighAnsi = "Times New Roman",
        };

        var fontSize = new FontSize
        {
            Val = "24",
        };

        runStyle.AppendChild(color);
        runStyle.AppendChild(font);
        runStyle.AppendChild(fontSize);
        blueTagStyle.AppendChild(runStyle);

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
            EastAsia = "Times New Roman",
            ComplexScript = "Times New Roman",
            HighAnsi = "Times New Roman",
        };

        var fontSize2 = new FontSize
        {
            Val = "24",
        };

        plainTextStyle.AppendChild(styleName2);
        plainTextStyle.AppendChild(linkedStyle2);
        var runStyle2 = new StyleRunProperties();
        runStyle2.AppendChild(color2);
        runStyle2.AppendChild(font2);
        runStyle2.AppendChild(fontSize2);
        plainTextStyle.AppendChild(runStyle2);

        #endregion

        #region Numbering 4

        var lastAbstractNumber = numStyles.Elements<AbstractNum>().Count() + 1;

        var lastNumNumber = numStyles.Elements<NumberingInstance>().Count() + 1;

        var numberingInstance = new NumberingInstance(new AbstractNumId
            {Val = lastAbstractNumber}) {NumberID = lastNumNumber};

        var abstractNumbering = new AbstractNum(
            new MultiLevelType
            {
                Val = new EnumValue<MultiLevelValues>(MultiLevelValues.Multilevel)
            },
            new Level
            {
                LevelIndex = 0,
                StartNumberingValue = new StartNumberingValue
                {
                    Val = 4
                },
                NumberingFormat = new NumberingFormat
                {
                    Val = new EnumValue<NumberFormatValues>(NumberFormatValues.Decimal)
                },
                LevelText = new LevelText
                {
                    Val = "%1"
                },
                LevelJustification = new LevelJustification
                {
                    Val = new EnumValue<LevelJustificationValues>(LevelJustificationValues.Left)
                },
                NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new Position
                    {
                        Val = "0"
                    },
                    new RightToLeftText
                    {
                        Val = OnOffValue.FromBoolean(false)
                    }),
            },
            new Level
            {
                LevelIndex = 1,
                StartNumberingValue = new StartNumberingValue
                {
                    Val = 1
                },
                NumberingFormat = new NumberingFormat
                {
                    Val = new EnumValue<NumberFormatValues>(NumberFormatValues.Decimal)
                },
                LevelText = new LevelText
                {
                    Val = "4.%2"
                },
                LevelJustification = new LevelJustification
                {
                    Val = new EnumValue<LevelJustificationValues>(LevelJustificationValues.Left)
                },
                NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new Position
                    {
                        Val = "0"
                    },
                    new RightToLeftText
                    {
                        Val = OnOffValue.FromBoolean(false)
                    })
            },
            new Level
            {
                LevelIndex = 2,
                StartNumberingValue = new StartNumberingValue
                {
                    Val = 1
                },
                NumberingFormat = new NumberingFormat
                {
                    Val = new EnumValue<NumberFormatValues>(NumberFormatValues.Decimal)
                },
                LevelText = new LevelText
                {
                    Val = "4.3.%3"
                },
                LevelJustification = new LevelJustification
                {
                    Val = new EnumValue<LevelJustificationValues>(LevelJustificationValues.Left)
                },
                NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new Position
                    {
                        Val = "0"
                    },
                    new RightToLeftText
                    {
                        Val = OnOffValue.FromBoolean(false)
                    })
            }
        ) {AbstractNumberId = lastAbstractNumber};

        var paragraphNumberStyle0 = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Numbering_4",
            CustomStyle = true,
            SemiHidden = new SemiHidden {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.Off)},
            StyleHidden = new StyleHidden {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.Off)},
            StyleName = new StyleName {Val = "Numbering_4 Lvl"},
            StyleParagraphProperties = new StyleParagraphProperties(new KeepNext
                {
                    Val = OnOffValue.FromBoolean(true)
                },
                new Spacing
                {
                    Val = 240
                },
                new OutlineLevel
                {
                    Val = 0
                },
                new Justification
                    {Val = new EnumValue<JustificationValues>(JustificationValues.Both)},
                new NumberingProperties(new NumberingId
                        {Val = lastNumNumber},
                    new NumberingLevelReference
                        {Val = 0})),
            StyleRunProperties = new StyleRunProperties(new Bold {Val = OnOffValue.FromBoolean(true)},
                new BoldComplexScript {Val = OnOffValue.FromBoolean(true)},
                new ItalicComplexScript {Val = OnOffValue.FromBoolean(true)},
                new FontSize {Val = "24"},
                new FontSizeComplexScript {Val = "26"}),
            PrimaryStyle = new PrimaryStyle {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On)},
            UnhideWhenUsed = new UnhideWhenUsed {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On)},
            // BasedOn = new BasedOn {Val = "phbase"},
            // NextParagraphStyle = new NextParagraphStyle {Val = "6"},
        };

        var paragraphNumberStyle1 = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Numbering_4_x",
            CustomStyle = true,
            SemiHidden = new SemiHidden {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.Off)},
            StyleHidden = new StyleHidden {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.Off)},
            StyleName = new StyleName {Val = "Numbering_4_x Lvl"},
            StyleParagraphProperties = new StyleParagraphProperties(new KeepNext
                {
                    Val = OnOffValue.FromBoolean(true)
                },
                new Spacing
                {
                    Val = 240
                },
                new OutlineLevel
                {
                    Val = 1
                },
                new Justification
                    {Val = new EnumValue<JustificationValues>(JustificationValues.Both)},
                new NumberingProperties(new NumberingId
                        {Val = lastNumNumber},
                    new NumberingLevelReference
                        {Val = 1})),
            StyleRunProperties = new StyleRunProperties(new Bold {Val = OnOffValue.FromBoolean(true)},
                new BoldComplexScript {Val = OnOffValue.FromBoolean(true)},
                new ItalicComplexScript {Val = OnOffValue.FromBoolean(true)},
                new FontSize {Val = "24"},
                new FontSizeComplexScript {Val = "26"}),
            PrimaryStyle = new PrimaryStyle {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On)},
            UnhideWhenUsed = new UnhideWhenUsed {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On)},
            // BasedOn = new BasedOn {Val = "phbase"},
            // NextParagraphStyle = new NextParagraphStyle {Val = "6"},
        };

        var paragraphNumberStyle2 = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Numbering_4_3_x",
            CustomStyle = true,
            SemiHidden = new SemiHidden {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.Off)},
            StyleHidden = new StyleHidden {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.Off)},
            StyleName = new StyleName {Val = "Numbering_4_3_x Lvl"},
            StyleParagraphProperties = new StyleParagraphProperties(new KeepNext
                {
                    Val = OnOffValue.FromBoolean(true)
                },
                new Spacing
                {
                    Val = 240
                },
                new OutlineLevel
                {
                    Val = 2
                },
                new Justification
                    {Val = new EnumValue<JustificationValues>(JustificationValues.Both)},
                new NumberingProperties(new NumberingId
                        {Val = lastNumNumber},
                    new NumberingLevelReference
                        {Val = 2})),
            StyleRunProperties = new StyleRunProperties(new Bold {Val = OnOffValue.FromBoolean(true)},
                new BoldComplexScript {Val = OnOffValue.FromBoolean(true)},
                new ItalicComplexScript {Val = OnOffValue.FromBoolean(true)},
                new FontSize {Val = "24"},
                new FontSizeComplexScript {Val = "26"}),
            PrimaryStyle = new PrimaryStyle {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On)},
            UnhideWhenUsed = new UnhideWhenUsed {Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On)},
            // BasedOn = new BasedOn {Val = "phbase"},
            // NextParagraphStyle = new NextParagraphStyle {Val = "6"},
        };

        if (lastAbstractNumber == 1)
        {
            numStyles.Append(abstractNumbering);
        }
        else
        {
            numStyles.Elements<AbstractNum>().Last().InsertAfterSelf(abstractNumbering);
        }

        if (lastNumNumber == 1)
        {
            numStyles.Append(numberingInstance);
        }
        else
        {
            numStyles.Elements<NumberingInstance>().Last().InsertAfterSelf(numberingInstance);
        }

        #endregion

        styles?.AppendChild(plainTextStyle);
        styles?.AppendChild(blueTagStyle);
        styles?.AppendChild(paragraphNumberStyle0);
        styles?.AppendChild(paragraphNumberStyle1);
        styles?.AppendChild(paragraphNumberStyle2);

        return (lastAbstractNumber, lastNumNumber);
    }

    private void SetXml(Paragraph p, IEnumerable<XElement> elements, int level, bool localName = true)
    {
        foreach (var elem in elements)
        {
            var indent = level * 2;
            var newtext = new Text();
            var newrun = new Run();
            var runProp = newrun.RunProperties ?? (newrun.RunProperties = new RunProperties());
            var elementName = localName ? elem.Name.LocalName : "xs:" + elem.Name.LocalName;

            runProp.RunStyle = new RunStyle
            {
                Val = "blueTag",
            };

            newtext.Space = SpaceProcessingModeValues.Preserve;

            newtext.Text =
                $"{new string(' ', indent)}<{elementName}{(elem.Attributes().Any() ? " " : "")}{string.Join(" ", elem.Attributes())}";

            if (level > 0)
            {
                newrun.Append(new Break());
            }

            newrun.Append(newtext);
            p.Append(newrun);

            SetXml(p, elem.Elements(), level + 1);

            //если есть внутри элементы
            if (elem.HasElements)
            {
                newtext.Text += ">";
            }
            //если внутри нет элементов и есть текст
            else if (!elem.HasElements && !string.IsNullOrWhiteSpace(elem.Value))
            {
                newtext.Text += ">";
                var innerTextArray = elem.Value.Split('\n');

                for (var i = 0; i < innerTextArray.Length; i++)
                {
                    var textPart = TrimAllChars(innerTextArray[i], ' ', '\r', '\n');
                    newtext = new Text();
                    newrun = new Run();
                    runProp = newrun.RunProperties ?? (newrun.RunProperties = new RunProperties());

                    runProp.RunStyle = new RunStyle
                    {
                        Val = "plainText",
                    };

                    newtext.Text = textPart;
                    newrun.Append(newtext);

                    if (i != innerTextArray.Length - 1 && !string.IsNullOrWhiteSpace(textPart))
                    {
                        newrun.Append(new Break());
                    }

                    p.Append(newrun);
                }

                if (innerTextArray.Length == 1)
                    indent = 0;
            }
            //если нет элементов и текста
            else if (!elem.HasElements && string.IsNullOrWhiteSpace(elem.Value))
            {
                newtext.Text += " />";

                continue;
            }

            newtext = new Text();
            newrun = new Run();
            runProp = newrun.RunProperties ?? (newrun.RunProperties = new RunProperties());

            newtext.Space = SpaceProcessingModeValues.Preserve;

            runProp.RunStyle = new RunStyle
            {
                Val = "blueTag",
            };

            newtext.Text = $"{new string(' ', indent)}</{elementName}>";

            if (elem.Elements().Any())
            {
                newrun.Append(new Break());
            }

            newrun.Append(newtext);
            p.Append(newrun);
        }
    }

    #endregion

    private string TrimAllChars(string text, params char[] replacedChars)
    {
        foreach (var replacedChar in replacedChars)
        {
            if (text.StartsWith(replacedChar) || text.EndsWith(replacedChar))
            {
                text = text.Trim(replacedChar);
                TrimAllChars(text, replacedChar);
            }
        }

        return text;
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

        using var newDoc = GetNewWordDocument();

        var (_, numberId) = InitStyles(newDoc);

        if (!string.IsNullOrEmpty(_entry.EtalonFolder) || !string.IsNullOrEmpty(_entry.TestFolder))
        {
            ExecuteInternal(newDoc);
        }

        if (!string.IsNullOrEmpty(_entry.TargetXsdFile))
        {
            ExecuteInternalXsd(newDoc, numberId);
        }

        if (!newDoc.AutoSave)
        {
            newDoc.Save();
        }
    }

    private WordprocessingDocument GetNewWordDocument()
    {
        var fileInfo = new FileInfo(_entry.TargetFile ?? throw new InvalidOperationException());

        var newFilePath =
            $"{_entry.SavePath}\\Готовый_{DateTime.Now.ToString("yyyy'_'MM'_'dd'-'HH'_'mm'_'ss")}_{fileInfo.Name.Split(".")[0]}{fileInfo.Extension}";

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