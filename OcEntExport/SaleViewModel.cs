using CommunityToolkit.Mvvm.ComponentModel;

namespace OcEntExport;

public partial class SaleViewModel : ObservableObject
{
    [ObservableProperty]
    private DateTime _businessDate;

    [ObservableProperty]
    private DateTime _checkOpen;

    [ObservableProperty]
    private DateTime _checkClose;

    [ObservableProperty]
    private long _checkId;

    [ObservableProperty]
    private long _checkDetailId;

    [ObservableProperty]
    private short _roundNumber;

    [ObservableProperty]
    private int _checkNumber;

    [ObservableProperty]
    private int _rvcNumber;

    [ObservableProperty]
    private string _rvcName = string.Empty;

    [ObservableProperty]
    private string _workstations = string.Empty;

    [ObservableProperty]
    private short _orderTypeIndex;

    [ObservableProperty]
    private string _orderTypeName = string.Empty;

    [ObservableProperty]
    private int _servingPeriodNumber;

    [ObservableProperty]
    private string _servingPeriodName = string.Empty;

    [ObservableProperty]
    private string _checkName = string.Empty;

    [ObservableProperty]
    private string _diningTableName = string.Empty;

    [ObservableProperty]
    private int _covers;

    [ObservableProperty]
    private string _payment = string.Empty;

    [ObservableProperty]
    private string _paymentNotes = string.Empty;

    [ObservableProperty]
    private long _miNumber;

    [ObservableProperty]
    private string _miMasterName = string.Empty;

    [ObservableProperty]
    private string _miDefName = string.Empty;

    [ObservableProperty]
    private long _familyGroupNum;

    [ObservableProperty]
    private string _familyGroupName = string.Empty;

    [ObservableProperty]
    private long _majorGroupNum;

    [ObservableProperty]
    private string _majorGroupName = string.Empty;

    [ObservableProperty]
    private long _salesCount;

    [ObservableProperty]
    private decimal _salesTotal;

    [ObservableProperty]
    private decimal _costTotal;

    [ObservableProperty]
    private decimal _discountTotal;

    [ObservableProperty]
    private decimal _salesAfterDiscount;

    [ObservableProperty]
    private string _discountNames = string.Empty;

    [ObservableProperty]
    private string _discountNotes = string.Empty;

    [ObservableProperty]
    private decimal _tax;

    [ObservableProperty]
    private decimal _serviceCharge;

    [ObservableProperty]
    private int _micNumber;

    [ObservableProperty]
    private string _micName = string.Empty;

    [ObservableProperty]
    private long _printClassNum;

    [ObservableProperty]
    private string _printClassName = string.Empty;

    [ObservableProperty]
    private int _empNumber;

    [ObservableProperty]
    private string _empName = string.Empty;

    [ObservableProperty]
    private int _authEmpNumber;

    [ObservableProperty]
    private string _authEmpName = string.Empty;

}