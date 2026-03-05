namespace BillMatch.Wpf.Models;

public class Transaction
{
    public DateTime Date { get; set; }
    public decimal Amount { get; set; }
    public string? CardNumber { get; set; }
    public string? Description { get; set; }
    public string? Account1 { get; set; }
    public string? Account2 { get; set; }
    public bool IsMatched { get; set; }
}
