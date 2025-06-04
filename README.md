# TaiwanLottery

This repository contains a C# console application that estimates Power Lottery (威力彩) number probabilities from historical data stored in Excel.

## Setup

1. Install the .NET 8 SDK or later.
2. Open `TaiwanLottery.sln` with Visual Studio or run commands via the solution.
3. Navigate to the `PowerLotteryPrediction` directory and restore dependencies:

```bash
dotnet restore
```

## Usage

Prepare an Excel file with at least 100 past lottery records. The file should
have the first sheet formatted as follows:

- Columns **A** to **F** contain the six main numbers.
- Column **G** (optional) contains the special number.
- Row 1 is a header row.

All main numbers must be between **1** and **38**, and special numbers (column G)
must be between **1** and **8**. The program validates these ranges and throws
an error if any value is outside them.

Run the application by supplying the Excel file path:

```bash
dotnet run --project PowerLotteryPrediction <path-to-excel-file>
```

The program lets you choose among five statistical methods — frequency,
recency weighted, last 30 draws frequency, last 10 draws frequency, or a
hybrid that averages frequency and recency. Probabilities for the six main
numbers (1–38) and the special number (1–8) are reported separately. After
displaying probabilities, the program also suggests a set of six main numbers
and one special number with the highest calculated probabilities. It also
prints the least likely numbers for reference.
