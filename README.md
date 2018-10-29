# AsposeExtensions
Extensions class to easy import and export(todo) excel.

1 - Decorate your model with AsposeExtensions SheetColumn attribute;<br />

```csharp
public class ContasAPagarSheetRow 
{
    [SheetColumn]           
    public DateTime? Data { get; set; }
    [SheetColumn(1)]        
    public string Historico { get; set; }
    [SheetColumn("Doc.")]   
    public int Doc { get; set; }
    [SheetColumn(3)]
    public decimal? Valor { get; set; }
    [SheetColumn("vcto.")]
    public DateTime DataDeVencimento { get; set; }
    [SheetColumn(5)]
    public string Pagto { get; set; }
    [SheetColumn(" Saldo ")]
    public string ValorFinal { get; set; }
}
```

1.1 - SheetColumn with empty constructor will search for column with same property name;<br />
1.2 - SheetColumn with number constructor will search for column with number of column;<br />
1.3 - SheetColumn with string constructor will search for column with name of string;<br />
1.3.1 - This feature is a method extension GetCellValueByNameColumn;<br />

<br />
2 - Call extension method RowsToModelList to read excel.

```csharp
  var excel = new Workbook();
  var rows = excel.RowsToModelList<ContasAPagarSheetRow>("SHEET NAME", 2);
```
