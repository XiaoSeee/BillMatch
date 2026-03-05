using BillMatch.Wpf.Models;
using BillMatch.Wpf.ViewModels;
using Xunit;

namespace BillMatch.Wpf.Tests;

public class MatchingAlgorithmTests
{
    [Fact]
    public void MatchTransactions_ShouldMatchExactAmountAndDate()
    {
        // Arrange
        var viewModel = new MainViewModel();
        var date = new DateTime(2023, 10, 1);
        
        var qianjiList = new List<Transaction>
        {
            new() { Date = date, Amount = -100.00m, Description = "Lunch" }
        };
        
        var billList = new List<Transaction>
        {
            new() { Date = date, Amount = -100.00m, Description = "Restaurant" }
        };

        // Act
        viewModel.MatchTransactions(qianjiList, billList);

        // Assert
        Assert.Single(viewModel.MatchedPairs);
        Assert.Empty(viewModel.UnmatchedQianji);
        Assert.Empty(viewModel.UnmatchedBills);
        Assert.True(qianjiList[0].IsMatched);
    }

    [Fact]
    public void MatchTransactions_ShouldMatchWithinDateTolerance()
    {
        // Arrange
        var viewModel = new MainViewModel { DaysTolerance = 2 };
        var billDate = new DateTime(2023, 10, 3);
        var qianjiDate = new DateTime(2023, 10, 1); // 2 days difference
        
        var qianjiList = new List<Transaction>
        {
            new() { Date = qianjiDate, Amount = -100.00m }
        };
        
        var billList = new List<Transaction>
        {
            new() { Date = billDate, Amount = -100.00m }
        };

        // Act
        viewModel.MatchTransactions(qianjiList, billList);

        // Assert
        Assert.Single(viewModel.MatchedPairs);
    }

    [Fact]
    public void MatchTransactions_ShouldNotMatchOutsideDateTolerance()
    {
        // Arrange
        var viewModel = new MainViewModel { DaysTolerance = 1 };
        var billDate = new DateTime(2023, 10, 3);
        var qianjiDate = new DateTime(2023, 10, 1); // 2 days difference
        
        var qianjiList = new List<Transaction>
        {
            new() { Date = qianjiDate, Amount = -100.00m }
        };
        
        var billList = new List<Transaction>
        {
            new() { Date = billDate, Amount = -100.00m }
        };

        // Act
        viewModel.MatchTransactions(qianjiList, billList);

        // Assert
        Assert.Empty(viewModel.MatchedPairs);
        Assert.Single(viewModel.UnmatchedQianji);
        Assert.Single(viewModel.UnmatchedBills);
    }

    [Fact]
    public void MatchTransactions_ShouldMatchAbsoluteAmount()
    {
        // Arrange
        var viewModel = new MainViewModel();
        var date = new DateTime(2023, 10, 1);
        
        var qianjiList = new List<Transaction>
        {
            new() { Date = date, Amount = 100.00m } // Positive in Qianji (Income or just different sign)
        };
        
        var billList = new List<Transaction>
        {
            new() { Date = date, Amount = -100.00m } // Negative in Bill
        };

        // Act
        viewModel.MatchTransactions(qianjiList, billList);

        // Assert
        Assert.Single(viewModel.MatchedPairs);
    }

    [Fact]
    public void MatchTransactions_ShouldHandleMultipleCandidates_PickClosestDate()
    {
        // Arrange
        var viewModel = new MainViewModel { DaysTolerance = 2 };
        var billDate = new DateTime(2023, 10, 3);
        
        var qianjiList = new List<Transaction>
        {
            new() { Date = new DateTime(2023, 10, 1), Amount = -100.00m, Description = "Farther" },
            new() { Date = new DateTime(2023, 10, 2), Amount = -100.00m, Description = "Closer" }
        };
        
        var billList = new List<Transaction>
        {
            new() { Date = billDate, Amount = -100.00m }
        };

        // Act
        viewModel.MatchTransactions(qianjiList, billList);

        // Assert
        Assert.Single(viewModel.MatchedPairs);
        var matched = viewModel.MatchedPairs[0];
        Assert.Equal("Closer", matched.QianjiTransaction.Description);
    }

    [Fact]
    public void MatchTransactions_ShouldNotDoubleMatch()
    {
        // Arrange
        var viewModel = new MainViewModel();
        var date = new DateTime(2023, 10, 1);
        
        var qianjiList = new List<Transaction>
        {
            new() { Date = date, Amount = -100.00m }
        };
        
        var billList = new List<Transaction>
        {
            new() { Date = date, Amount = -100.00m },
            new() { Date = date, Amount = -100.00m }
        };

        // Act
        viewModel.MatchTransactions(qianjiList, billList);

        // Assert
        Assert.Single(viewModel.MatchedPairs);
        Assert.Single(viewModel.UnmatchedBills);
    }
}
