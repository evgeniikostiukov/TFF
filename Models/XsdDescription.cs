namespace tff.main.Models;

/// <summary>
///     Состав передаваемой информации
/// </summary>
public class XsdDescription
{
    /// <summary>
    ///     Номер пункта
    /// </summary>
    public string Number { get; set; }

    /// <summary>
    ///     Код поля
    /// </summary>
    public string Field { get; set; }

    /// <summary>
    ///     Описание поля
    /// </summary>
    public string Annotation { get; set; }

    /// <summary>
    ///     Требование к заполнению
    /// </summary>
    public bool IsRequired { get; set; }

    /// <summary>
    ///     Способ заполнения/Тип
    /// </summary>
    public string TypeOrFillMethod { get; set; }

    /// <summary>
    ///     Комментарий
    /// </summary>
    public string Comment { get; set; }

    /// <summary>
    ///     Заголовок подраздела
    /// </summary>
    public string Title { get; set; }

    /// <summary>
    ///     Родитель
    /// </summary>
    public XsdDescription Parent { get; set; }
}