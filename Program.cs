using ExcelDataReader;
using System.Text;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

Console.WriteLine("=================================");
Console.WriteLine("Excel Compare Tool");
Console.WriteLine("=================================");

// =====================================================
// FILE PATHS
// =====================================================

string file1Path = "input/file1.xlsx";
string file2Path = "input/file2.xlsx";
string file3Path = "input/file3.xlsx";

// =====================================================
// CHECK REQUIRED FILES
// =====================================================

CheckFile(file1Path);
CheckFile(file2Path);

// FILE 3 OPTIONAL
bool hasFile3 = File.Exists(file3Path);

// =====================================================
// FILTER IDS
// =====================================================

var ids = new HashSet<string>
{
    "236100"
};

// =====================================================
// FILE CONFIGS
// =====================================================

// FILE 1
var file1Config = new FileConfig
{
    IdColumn = "MAINACCOUNTID",
    NumberColumn = "CONTRACTNUMBER",

    DebitColumn = "DEBIT",
    CreditColumn = "CREDIT"
};

// FILE 2
// 01 = DEBIT
// 02 = CREDIT 
var file2Config = new FileConfig
{
    IdColumn = "ACCOUNT_CODE",
    NumberColumn = "CONTRACT_NO",

    UseTransactionType = true,

    TransactionTypeColumn = "DEBIT_CREDIT_DIFF",

    AmountColumn = "TRADE_AMOUNT",

    DebitTypeValue = "01",
    CreditTypeValue = "02"
};

// FILE 3
var file3Config = new FileConfig
{
    IdColumn = "ACCID",
    NumberColumn = "CNTRNO",

    DebitColumn = "DEBITAMT",
    CreditColumn = "CREDITAMT"
};

// =====================================================
// BUILD SUMMARY
// =====================================================

Console.WriteLine("Reading File 1...");

var file1Summary = BuildSummary(
    file1Path,
    ids,
    file1Config
);

Console.WriteLine("Reading File 2...");

var file2Summary = BuildSummary(
    file2Path,
    ids,
    file2Config
);

// OPTIONAL FILE 3
var file3Summary =
    new Dictionary<(string Id, string Number), Summary>();

if (hasFile3)
{
    Console.WriteLine("Reading File 3...");

    file3Summary = BuildSummary(
        file3Path,
        ids,
        file3Config
    );
}
else
{
    Console.WriteLine("File 3 not found -> Skip");
}

// =====================================================
// UNION ALL KEYS
// =====================================================

var allKeys = file1Summary.Keys
    .Union(file2Summary.Keys);

if (hasFile3)
{
    allKeys = allKeys.Union(file3Summary.Keys);
}

var finalKeys = allKeys
    .Distinct()
    .OrderBy(x => x.Id)
    .ThenBy(x => x.Number)
    .ToList();

// =====================================================
// OUTPUT
// =====================================================

Directory.CreateDirectory("output");

var lines = new List<string>();

lines.Add(
    "AccountId," +
    "ContractNo," +

    // FILE 1
    "File1DebitSum," +
    "File1RecordDebit," +
    "File1CreditSum," +
    "File1RecordCredit," +
    "File1Amount," +

    // FILE 2
    "File2DebitSum," +
    "File2RecordDebit," +
    "File2CreditSum," +
    "File2RecordCredit," +
    "File2Amount," +

    // FILE 3
    "File3DebitSum," +
    "File3RecordDebit," +
    "File3CreditSum," +
    "File3RecordCredit," +
    "File3Amount"
);

// =====================================================
// COMPARE
// =====================================================

foreach (var key in finalKeys)
{
    bool has1 = file1Summary.ContainsKey(key);

    bool has2 = file2Summary.ContainsKey(key);

    bool has3 =
        hasFile3 &&
        file3Summary.ContainsKey(key);

    var data1 = has1
        ? file1Summary[key]
        : new Summary();

    var data2 = has2
        ? file2Summary[key]
        : new Summary();

    var data3 = has3
        ? file3Summary[key]
        : new Summary();

    bool sameDebit;

    if (hasFile3)
    {
        sameDebit =
            data1.DebitSum == data2.DebitSum &&
            data2.DebitSum == data3.DebitSum &&

            data1.RecordDebit == data2.RecordDebit &&
            data2.RecordDebit == data3.RecordDebit;
    }
    else
    {
        sameDebit =
            data1.DebitSum == data2.DebitSum &&
            data1.RecordDebit == data2.RecordDebit;
    }

    bool sameCredit;

    if (hasFile3)
    {
        sameCredit =
            data1.CreditSum == data2.CreditSum &&
            data2.CreditSum == data3.CreditSum &&

            data1.RecordCredit == data2.RecordCredit &&
            data2.RecordCredit == data3.RecordCredit;
    }
    else
    {
        sameCredit =
            data1.CreditSum == data2.CreditSum &&
            data1.RecordCredit == data2.RecordCredit;
    }

    // =================================================
    // AMOUNT
    // =================================================

    decimal file1Amount =
        data1.DebitSum - data1.CreditSum;

    decimal file2Amount =
        data2.DebitSum - data2.CreditSum;

    decimal file3Amount =
        data3.DebitSum - data3.CreditSum;

    // =================================================
    // EXPORT
    // =================================================

    lines.Add(
        $"{key.Id}," +
        $"{key.Number}," +

        // FILE 1
        $"{data1.DebitSum}," +
        $"{data1.RecordDebit}," +
        $"{data1.CreditSum}," +
        $"{data1.RecordCredit}," +
        $"{file1Amount}," +

        // FILE 2
        $"{data2.DebitSum}," +
        $"{data2.RecordDebit}," +
        $"{data2.CreditSum}," +
        $"{data2.RecordCredit}," +
        $"{file2Amount}," +

        // FILE 3
        $"{data3.DebitSum}," +
        $"{data3.RecordDebit}," +
        $"{data3.CreditSum}," +
        $"{data3.RecordCredit}," +
        $"{file3Amount}"
    );
}

// =====================================================
// WRITE CSV
// =====================================================

File.WriteAllLines(
    "output/compare_result.csv",
    lines
);

Console.WriteLine();
Console.WriteLine("=================================");
Console.WriteLine("DONE");
Console.WriteLine("=================================");
Console.WriteLine("Output: output/compare_result.csv");

Console.ReadLine();

// =====================================================
// BUILD SUMMARY
// =====================================================

static Dictionary<(string Id, string Number), Summary>
BuildSummary(
    string filePath,
    HashSet<string> ids,
    FileConfig config
)
{
    var summaryMap =
        new Dictionary<(string Id, string Number), Summary>();

    using var stream = File.Open(
        filePath,
        FileMode.Open,
        FileAccess.Read,
        FileShare.ReadWrite
    );

    using var reader =
        ExcelReaderFactory.CreateReader(stream);

    // =================================================
    // READ HEADER
    // =================================================

    if (!reader.Read())
    {
        return summaryMap;
    }

    var columnMap =
        new Dictionary<string, int>();

    for (int i = 0; i < reader.FieldCount; i++)
    {
        var columnName = reader.GetValue(i)?
            .ToString()?
            .Trim()?
            .ToUpper();

        if (!string.IsNullOrEmpty(columnName))
        {
            columnMap[columnName] = i;
        }
    }

    // =================================================
    // REQUIRED COLUMNS
    // =================================================

    var requiredColumns = new List<string>
    {
        config.IdColumn.ToUpper(),
        config.NumberColumn.ToUpper()
    };

    if (config.UseTransactionType)
    {
        requiredColumns.Add(
            config.TransactionTypeColumn.ToUpper()
        );

        requiredColumns.Add(
            config.AmountColumn.ToUpper()
        );
    }
    else
    {
        requiredColumns.Add(
            config.DebitColumn.ToUpper()
        );

        requiredColumns.Add(
            config.CreditColumn.ToUpper()
        );
    }

    foreach (var col in requiredColumns)
    {
        if (!columnMap.ContainsKey(col))
        {
            throw new Exception(
                $"Column not found -> {col}"
            );
        }
    }

    // =================================================
    // GET COLUMN INDEX
    // =================================================

    int idIndex =
        columnMap[config.IdColumn.ToUpper()];

    int numberIndex =
        columnMap[config.NumberColumn.ToUpper()];

    int debitIndex = -1;

    int creditIndex = -1;

    int typeIndex = -1;

    int amountIndex = -1;

    if (config.UseTransactionType)
    {
        typeIndex =
            columnMap[
                config.TransactionTypeColumn.ToUpper()
            ];

        amountIndex =
            columnMap[
                config.AmountColumn.ToUpper()
            ];
    }
    else
    {
        debitIndex =
            columnMap[
                config.DebitColumn.ToUpper()
            ];

        creditIndex =
            columnMap[
                config.CreditColumn.ToUpper()
            ];
    }

    // =================================================
    // READ DATA
    // =================================================

    while (reader.Read())
    {
        try
        {
            // =============================================
            // ID
            // =============================================

            var id = reader.GetValue(idIndex)?
                .ToString()?
                .Trim();

            if (string.IsNullOrEmpty(id))
                continue;

            // FILTER IDS
            if (!ids.Contains(id))
                continue;

            // =============================================
            // NUMBER
            // =============================================

            var number = reader.GetValue(numberIndex)?
                .ToString()?
                .Trim();

            if (string.IsNullOrEmpty(number))
            {
                number = "UNKNOWN";
            }

            var key = (id, number);

            // CREATE GROUP
            if (!summaryMap.ContainsKey(key))
            {
                summaryMap[key] = new Summary();
            }

            // =============================================
            // TYPE MODE
            // FILE 2
            // =============================================

            if (config.UseTransactionType)
            {
                var transactionType =
                    reader.GetValue(typeIndex)?
                    .ToString()?
                    .Trim();

                decimal amount =
                    ParseDecimal(
                        reader.GetValue(amountIndex)
                    );

                // DEBUG
                // Console.WriteLine(
                //     $"TYPE={transactionType} AMOUNT={amount}"
                // );

                // DEBIT
                if (
                    transactionType ==
                    config.DebitTypeValue
                )
                {
                    summaryMap[key].DebitSum += amount;

                    summaryMap[key].RecordDebit++;
                }

                // CREDIT
                else if (
                    transactionType ==
                    config.CreditTypeValue
                )
                {
                    summaryMap[key].CreditSum += amount;

                    summaryMap[key].RecordCredit++;
                }
            }

            // =============================================
            // NORMAL MODE
            // =============================================

            else
            {
                decimal debit =
                    ParseDecimal(
                        reader.GetValue(debitIndex)
                    );

                decimal credit =
                    ParseDecimal(
                        reader.GetValue(creditIndex)
                    );

                // DEBIT
                if (debit != 0)
                {
                    summaryMap[key].DebitSum += debit;

                    summaryMap[key].RecordDebit++;
                }

                // CREDIT
                if (credit != 0)
                {
                    summaryMap[key].CreditSum += credit;

                    summaryMap[key].RecordCredit++;
                }
            }
        }
        catch
        {
            // SKIP INVALID ROW
        }
    }

    return summaryMap;
}

// =====================================================
// PARSE DECIMAL
// =====================================================

static decimal ParseDecimal(object? value)
{
    if (value == null)
        return 0;

    var text = value
        .ToString()?
        .Trim()
        .Replace(",", "");

    if (string.IsNullOrEmpty(text))
        return 0;

    decimal.TryParse(
        text,
        out decimal result
    );

    return result;
}

// =====================================================
// CHECK FILE
// =====================================================

static void CheckFile(string filePath)
{
    if (!File.Exists(filePath))
    {
        Console.WriteLine(
            $"ERROR: File not found -> {filePath}"
        );

        Console.ReadLine();

        Environment.Exit(0);
    }
}

// =====================================================
// FILE CONFIG
// =====================================================

public class FileConfig
{
    public string IdColumn { get; set; } = "";

    public string NumberColumn { get; set; } = "";

    // NORMAL MODE
    public string DebitColumn { get; set; } = "";

    public string CreditColumn { get; set; } = "";

    // TYPE MODE
    public bool UseTransactionType { get; set; }

    public string TransactionTypeColumn { get; set; } = "";

    public string AmountColumn { get; set; } = "";

    public string DebitTypeValue { get; set; } = "";

    public string CreditTypeValue { get; set; } = "";
}

// =====================================================
// SUMMARY
// =====================================================

public class Summary
{
    // DEBIT
    public decimal DebitSum { get; set; }

    public int RecordDebit { get; set; }

    // CREDIT
    public decimal CreditSum { get; set; }

    public int RecordCredit { get; set; }
}
