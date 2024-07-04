using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;

namespace OcEntExport;

public partial class SummaryL1 : ObservableObject
{
    public List<string> DistinctMajorGroups { get; }

    [ObservableProperty]
    private string _type = string.Empty;

    [ObservableProperty]
    public ObservableCollection<SummaryL2> _result = new ObservableCollection<SummaryL2>();

    public SummaryL1(List<string> distinctMajorGroups)
    {
        DistinctMajorGroups = distinctMajorGroups;
    }

    public decimal GetSalesTotal(string majorGroup) => Result.Sum(r => r.GetSalesTotal(majorGroup));
    public decimal GetCostTotal(string majorGroup) => Result.Sum(r => r.GetCostTotal(majorGroup));

    public decimal GrandSalesTotal { get => Result.Sum(s => s.GrandSalesTotal); }
    public decimal GrandCostTotal { get => Result.Sum(s => s.GrandCostTotal); }
}

public partial class SummaryL2 : ObservableObject
{
    [ObservableProperty]
    private string _paymentNotes = string.Empty;

    [ObservableProperty]
    public ObservableCollection<SummaryL3> _result = new ObservableCollection<SummaryL3>();

    public decimal GetSalesTotal(string majorGroup) => Result.Sum(r => r.GetSalesTotal(majorGroup));
    public decimal GetCostTotal(string majorGroup) => Result.Sum(r => r.GetCostTotal(majorGroup));

    public decimal GrandSalesTotal { get => Result.Sum(s => s.GrandSalesTotal); }
    public decimal GrandCostTotal { get => Result.Sum(s => s.GrandCostTotal); }
}

public partial class SummaryL3 : ObservableObject
{
    [ObservableProperty]
    private string _rvcName = string.Empty;

    [ObservableProperty]
    public ObservableCollection<SaleViewModel> _result = new ObservableCollection<SaleViewModel>();

    public decimal GetSalesTotal(string majorGroup) => Result.Where(s => s.MajorGroupName == majorGroup).Sum(s => s.SalesTotal);
    public decimal GetCostTotal(string majorGroup) => Result.Where(s => s.MajorGroupName == majorGroup).Sum(s => s.CostTotal);

    public decimal GrandSalesTotal { get => Result.Sum(s => s.SalesTotal); }
    public decimal GrandCostTotal { get => Result.Sum(s => s.CostTotal); }
}
