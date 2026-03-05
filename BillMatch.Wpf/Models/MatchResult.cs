namespace BillMatch.Wpf.Models;

public class MatchResult
{
    public Transaction BillTransaction { get; set; } = null!;
    public Transaction? QianjiTransaction { get; set; }
    public string Status { get; set; } = null!;
}
