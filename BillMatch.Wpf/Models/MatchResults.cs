namespace BillMatch.Wpf.Models;

public class MatchResults
{
    public List<Transaction> UnmatchedBills { get; set; } = new();
    public List<MatchResult> MatchedPairs { get; set; } = new();
    public List<Transaction> UnmatchedQianji { get; set; } = new();
}
