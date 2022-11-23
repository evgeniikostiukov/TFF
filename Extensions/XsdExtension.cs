using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using tff.main.Models;

namespace tff.main.Extensions;

public static class XsdExtension
{
    public static void InsertTableHeader(this Table table)
    {
        var tableRowProperties =
            new TableRowProperties(new CantSplit
                {
                    Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On),
                },
                new TableHeader
                {
                    Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On),
                }
            );

        var headerRow = new TableRow
        {
            TableRowProperties = tableRowProperties,
        };

        var cell1 = new TableCell();
        var cell2 = new TableCell();
        var cell3 = new TableCell();
        var cell4 = new TableCell();
        var cell5 = new TableCell();
        var cell6 = new TableCell();

        var headerParagraphProperties1 = new ParagraphProperties(new ParagraphStyleId
            {
                Val = "af3",
            }
        );

        var headerParagraphProperties2 = new ParagraphProperties(new ParagraphStyleId
            {
                Val = "af3",
            }
        );

        var headerParagraphProperties3 = new ParagraphProperties(new ParagraphStyleId
            {
                Val = "af3",
            }
        );

        var headerParagraphProperties4 = new ParagraphProperties(new ParagraphStyleId
            {
                Val = "af3",
            }
        );

        var headerParagraphProperties5 = new ParagraphProperties(new ParagraphStyleId
            {
                Val = "af3",
            }
        );

        var headerParagraphProperties6 = new ParagraphProperties(new ParagraphStyleId
            {
                Val = "af3",
            }
        );

        var headerPara1 = new Paragraph(headerParagraphProperties1);
        var headerPara2 = new Paragraph(headerParagraphProperties2);
        var headerPara3 = new Paragraph(headerParagraphProperties3);
        var headerPara4 = new Paragraph(headerParagraphProperties4);
        var headerPara5 = new Paragraph(headerParagraphProperties5);
        var headerPara6 = new Paragraph(headerParagraphProperties6);

        var headerRun1 = new Run();
        var headerRun2 = new Run();
        var headerRun3 = new Run();
        var headerRun4 = new Run();
        var headerRun5 = new Run();
        var headerRun6 = new Run();

        var headerText1 = new Text();
        var headerText2 = new Text();
        var headerText3 = new Text();
        var headerText4 = new Text();
        var headerText5 = new Text();
        var headerText6 = new Text();

        headerText1.Text = "№ п/п";
        headerText2.Text = "Код поля";
        headerText3.Text = "Описание поля";
        headerText4.Text = "Требования к заполнению";
        headerText5.Text = "Способ заполнения/Тип";
        headerText6.Text = "Комментарий";

        headerRun1.AddElementChild(headerText1);
        headerRun2.AddElementChild(headerText2);
        headerRun3.AddElementChild(headerText3);
        headerRun4.AddElementChild(headerText4);
        headerRun5.AddElementChild(headerText5);
        headerRun6.AddElementChild(headerText6);

        headerPara1.AddElementChild(headerRun1);
        headerPara2.AddElementChild(headerRun2);
        headerPara3.AddElementChild(headerRun3);
        headerPara4.AddElementChild(headerRun4);
        headerPara5.AddElementChild(headerRun5);
        headerPara6.AddElementChild(headerRun6);

        cell1.AddElementChild(headerPara1);
        cell2.AddElementChild(headerPara2);
        cell3.AddElementChild(headerPara3);
        cell4.AddElementChild(headerPara4);
        cell5.AddElementChild(headerPara5);
        cell6.AddElementChild(headerPara6);

        headerRow.AddElementChild(cell1);
        headerRow.AddElementChild(cell2);
        headerRow.AddElementChild(cell3);
        headerRow.AddElementChild(cell4);
        headerRow.AddElementChild(cell5);
        headerRow.AddElementChild(cell6);

        table.AddElementChild(headerRow);
    }

    public static void InsertTableRow(this Table table,
        List<KeyValuePair<Paragraph, XsdDescription>> paragraphsWithHyperlink,
        XsdDescription[] elements,
        XsdDescription xsdDescription,
        int counter)
    {
        var tableRow = new TableRow();

        var paragraphProperties1 = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId
            {
                Val = "310",
            },
        };

        var paragraphProperties2 = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId
            {
                Val = "310",
            },
        };

        var paragraphProperties3 = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId
            {
                Val = "310",
            },
        };

        var paragraphProperties4 = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId
            {
                Val = "310",
            },
        };

        var paragraphProperties5 = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId
            {
                Val = "310",
            },
        };

        var paragraphProperties6 = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId
            {
                Val = "310",
            },
        };

        var tablePara1 = new Paragraph(paragraphProperties1);
        var tablePara2 = new Paragraph(paragraphProperties2);
        var tablePara3 = new Paragraph(paragraphProperties3);
        var tablePara4 = new Paragraph(paragraphProperties4);
        var tablePara5 = new Paragraph(paragraphProperties5);
        var tablePara6 = new Paragraph(paragraphProperties6);

        var tableRun1 = new Run();
        var tableRun2 = new Run();
        var tableRun3 = new Run();
        var tableRun4 = new Run();
        var tableRun5 = new Run();
        var tableRun6 = new Run();

        var cell1 = new TableCell();
        var cell2 = new TableCell();
        var cell3 = new TableCell();
        var cell4 = new TableCell();
        var cell5 = new TableCell();
        var cell6 = new TableCell();

        var cellProperties1 = new TableCellProperties(new TableCellWidth
            {
                Type = TableWidthUnitValues.Pct,
                Width = "7",
            },
            new NoWrap(),
            new HideMark(),
            new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "FFFFFF",
            }
        );

        var cellProperties2 = new TableCellProperties(new TableCellWidth
            {
                Type = TableWidthUnitValues.Pct,
                Width = "18",
            },
            new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "FFFFFF",
            }
        );

        var cellProperties3 = new TableCellProperties(new TableCellWidth
            {
                Type = TableWidthUnitValues.Pct,
                Width = "20",
            },
            new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "FFFFFF",
            }
        );

        var cellProperties4 = new TableCellProperties(new TableCellWidth
            {
                Type = TableWidthUnitValues.Pct,
                Width = "20",
            },
            new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "FFFFFF",
            }
        );

        var cellProperties5 = new TableCellProperties(new TableCellWidth
            {
                Type = TableWidthUnitValues.Pct,
                Width = "15",
            },
            new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "FFFFFF",
            }
        );

        var cellProperties6 = new TableCellProperties(new TableCellWidth
            {
                Type = TableWidthUnitValues.Pct,
                Width = "20",
            },
            new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "FFFFFF",
            }
        );

        cell1.AppendChild(cellProperties1);
        cell2.AppendChild(cellProperties2);
        cell3.AppendChild(cellProperties3);
        cell4.AppendChild(cellProperties4);
        cell5.AppendChild(cellProperties5);
        cell6.AppendChild(cellProperties6);

        tablePara1.AddElementChild(tableRun1);
        tablePara2.AddElementChild(tableRun2);
        tablePara3.AddElementChild(tableRun3);
        tablePara4.AddElementChild(tableRun4);
        tablePara5.AddElementChild(tableRun5);
        tablePara6.AddElementChild(tableRun6);

        if (!string.IsNullOrEmpty(xsdDescription.TypeOrFillMethod)
         && xsdDescription.Parent != null
         && !xsdDescription.TypeOrFillMethod.StartsWith("xs:"))
        {
            var (fieldName, _) = GetFieldNameAndMethod(xsdDescription);

            var reference = WalkFindReference(elements, xsdDescription.Parent, fieldName);

            if (reference != null)
            {
                paragraphsWithHyperlink.Add(new KeyValuePair<Paragraph, XsdDescription>(tablePara6, reference));
            }
        }

        tableRun1.SetRunTextWithNewLines(counter.ToString());
        tableRun2.SetRunTextWithNewLines(xsdDescription.Field);
        tableRun3.SetRunTextWithNewLines(xsdDescription.Annotation);

        tableRun4.SetRunTextWithNewLines($"{(xsdDescription.IsRequired ? "Обязательно" : "Не обязательно")} к заполнению");

        tableRun5.SetRunTextWithNewLines(xsdDescription.TypeOrFillMethod);
        tableRun6.SetRunTextWithNewLines(xsdDescription.Comment);

        cell1.AddElementChild(tablePara1);
        cell2.AddElementChild(tablePara2);
        cell3.AddElementChild(tablePara3);
        cell4.AddElementChild(tablePara4);
        cell5.AddElementChild(tablePara5);
        cell6.AddElementChild(tablePara6);

        tableRow.AppendChild(cell1);
        tableRow.AppendChild(cell2);
        tableRow.AppendChild(cell3);
        tableRow.AppendChild(cell4);
        tableRow.AppendChild(cell5);
        tableRow.AppendChild(cell6);

        table.AppendChild(tableRow);
    }

    public static Paragraph InsertTitle(OpenXmlElement nextElement,
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

    private static void AddElementChild<T>(this OpenXmlCompositeElement element, T child)
        where T : OpenXmlElement
    {
        element.AppendChild(child);
    }

    public static void AddBookmark(this OpenXmlElement element, string number, int bookmarkCount = 1)
    {
        for (var i = 0; i < bookmarkCount; i++)
        {
            var bookmarkStart = new BookmarkStart
            {
                Id = number,
                Name = number,
            };

            var bookmarkEnd = new BookmarkEnd
            {
                Id = number,
            };

            element.FirstChild?.InsertAfterSelf(bookmarkStart);
            element.LastChild?.InsertAfterSelf(bookmarkEnd);
        }
    }

    public static void AddHyperlink(Run run, string referenceNumber)
    {
        var runBegin = run.InsertAfterSelf(new Run(new FieldChar
                {
                    FieldCharType = FieldCharValues.Begin,
                }
            )
        );

        var runRef = runBegin.InsertAfterSelf(new Run(new FieldCode($" REF {referenceNumber} \\r \\h ")
                {
                    Space = SpaceProcessingModeValues.Preserve,
                }
            )
        );

        var runSeparate = runRef.InsertAfterSelf(new Run(new FieldChar
                {
                    FieldCharType = FieldCharValues.Separate,
                }
            )
        );

        var runText = runSeparate.InsertAfterSelf(new Run(new Text(referenceNumber)));

        var runEnd = runText.InsertAfterSelf(new Run(new FieldChar
                {
                    FieldCharType = FieldCharValues.End,
                }
            )
        );

        runEnd.InsertAfterSelf(new Run(new Text(")")));
    }

    private static void SetRunTextWithNewLines(this Run run, string text)
    {
        var newLineArray = new[] {Environment.NewLine, "\n", "\r\n", "\n\r",};
        var textArray = text.Split(newLineArray, StringSplitOptions.None);
        var first = true;

        foreach (var line in textArray)
        {
            if (!first)
            {
                run.Append(new Break());
            }

            first = false;

            var txt = new Text
            {
                Text = line,
                Space = SpaceProcessingModeValues.Preserve,
            };

            run.Append(txt);
        }
    }

    public static void SetReferenceComment(XsdDescription[] elements)
    {
        var childs = elements.Where(x => x.Parent != null || x.TypeOrFillMethod == "Составной элемент");

        //childrens
        foreach (var element in elements.Where(x => childs.Any(s => x == s)))
        {
            var (fieldName, method) = GetFieldNameAndMethod(element);

            var referenceElement = WalkFindReference(elements, element.Parent, fieldName);

            if (referenceElement == null)
            {
                continue;
            }

            element.Comment = $"{method} атрибут(см. описание в пункте ";
        }
    }

    private static XsdDescription WalkFindReference(XsdDescription[] elements, XsdDescription parent, string findName)
    {
        var reference = elements.FirstOrDefault(x => x.Parent == parent && x.Field == findName);

        if (reference != null)
        {
            return reference;
        }

        return parent != null ? WalkFindReference(elements, parent.Parent, findName) : null;
    }

    public static string GetAnnotation(XmlSchemaAnnotated schemaObject)
    {
        var documentationTag = schemaObject.Annotation?.Items[0] as XmlSchemaDocumentation;
        var markup = documentationTag?.Markup?.Length > 0 ? documentationTag.Markup[0]?.Value?.TrimEnd(':') : string.Empty;

        return markup;
    }

    private static (string fieldName, string method) GetFieldNameAndMethod(XsdDescription element)
    {
        string fieldName;
        string method;

        switch (element.TypeOrFillMethod)
        {
            case "Составной элемент":
                fieldName = element.Field;
                method = "Составной";

                break;
            default:
                fieldName = $"{element.TypeOrFillMethod}";
                method = "Множественный";

                break;
        }

        return (fieldName, method);
    }
}