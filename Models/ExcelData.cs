using System;

namespace ExcelProcessor.Models;

public class ExcelData
{
    public int RowNumber { get; set; }
    public decimal FirstNum { get; set; }
    public decimal SecondNum { get; set; }
    public decimal Result { get; set; }
    public string? Operation { get; set; }
}
