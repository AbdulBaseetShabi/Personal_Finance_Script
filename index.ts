import { existsSync, readdirSync } from "fs";
import {
  TableColumnProperties,
  Workbook,
  Worksheet,
  Style,
  Border,
} from "exceljs";
const open = require("open");
const dotenv = require("dotenv");
dotenv.config();

const CURRENT_YEAR = new Date().getFullYear();
const BUDGET_FILE = `./Budget/budget-${CURRENT_YEAR}.xlsx`;
const IGNORED_TRANSACTIONS = process.env.IGNORED_TRANSACTIONS
  ? process.env.IGNORED_TRANSACTIONS.split(";")
  : []; //used to ignore transfers between accounts - key has private info

console.log(IGNORED_TRANSACTIONS);
const TEMPLATE_FILE = `./Budget/template.xlsx`;

const EXPENSE_TYPES = ["Entertainment", "Utilities", "Others"];
const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const BLACK = { argb: "FF000000" };
const WHITE = { argb: "FFFFFFFF" };
const PURPLE = { argb: "FF7030A0" };
const GREEN = { argb: "FF9BBB59" };
const HEADER_STYLE: Partial<Style> = {
  alignment: {
    horizontal: "center",
  },
  font: {
    color: WHITE,
    bold: true,
    name: "Arial Narrow",
  },
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: PURPLE,
  },
};

const MAIN_HEADER_STYLE: Partial<Style> = {
  alignment: {
    horizontal: "center",
  },
  font: { color: WHITE, bold: true, name: "Arial Narrow", size: 16 },
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: BLACK,
  },
};
interface Budget {
  [key: string]: {
    budget: number;
    name: string;
    expenseType: string;
  };
}

interface AmountData {
  date: string;
  amount: number;
}

interface BankData {
  spending: {
    [key: string]: {
      totalAmount: number;
      data: AmountData[];
    };
  };
  income: {
    [key: string]: {
      totalAmount: number;
      data: AmountData[];
    };
  };

  totalSpending: number;
  totalIncome: number;
}

interface ExpenseSummaryData {
  name: string;
  type: string;
  budget: number;
  cost: number;
  diff: number;
}

// create kvp from template - has the latest budget data
const getBudget = async (): Promise<Budget | null> => {
  if (!existsSync(TEMPLATE_FILE)) {
    return null;
  }

  const templateWorkbook = new Workbook();
  await templateWorkbook.xlsx.readFile(TEMPLATE_FILE);

  const budgetSheet = templateWorkbook.getWorksheet("Budget");

  const parsedBudget: Budget = {};
  let lastExpenseCategory: string;

  budgetSheet.eachRow({ includeEmpty: false }, (row, _) => {
    if (!row) {
      return;
    }

    const name = row.getCell("A").value as string;
    const budget =
      row.getCell("B").value === null ? 0 : (row.getCell("B").value as number);
    const key = row.getCell("C").value as string;

    if (!key || !name) {
      console.log(`Key/Name not found for Name/Key: ${name} || ${key}`);
      return;
    }

    const lowerCaseName = name.toLowerCase();
    if (
      EXPENSE_TYPES.find((a) => a.toLowerCase() === lowerCaseName) ||
      lowerCaseName === "expense"
    ) {
      if (lowerCaseName !== "expense") {
        lastExpenseCategory =
          lowerCaseName[0].toUpperCase() + lowerCaseName.slice(1);
      }
      return;
    }

    parsedBudget[key] = {
      budget,
      name,
      expenseType: lastExpenseCategory,
    };
  });

  return parsedBudget;
};

const getDataFromArray = (array: string[], key: string): string | undefined =>
  array.find((element) => key.toLowerCase().includes(element.toLowerCase()));

// parse data from csv file - saving and chequing (group data by key - ignoring IGNORED_TRANSACTIONS)

// TODO: Mutliple files at once seperated by dates
const getBankData = async (
  filePath: string,
  budgetKeys: string[]
): Promise<BankData | null> => {
  if (!existsSync(filePath)) {
    return null;
  }

  const parsedBankData: BankData = {
    spending: {},
    income: {},
    totalSpending: 0,
    totalIncome: 0,
  };

  for (const file of readdirSync(filePath)) {
    if (!file.endsWith(".csv")) {
      continue;
    }

    const templateWorkbook = new Workbook();
    const csvFile = await templateWorkbook.csv.readFile(`${filePath}/${file}`);
    csvFile.eachRow({ includeEmpty: false }, (row, rowIndex) => {
      const firstColumn = (row.getCell("A").value as string).toLowerCase();
      if (rowIndex < 5 || firstColumn === "first bank card") {
        return;
      }

      const date = (row.getCell("C").value as number).toString();
      const amount = row.getCell("D").value as number;
      const key_description = row.getCell("E").value as string;

      if (getDataFromArray(IGNORED_TRANSACTIONS, key_description)) {
        return;
      }

      const key =
        getDataFromArray(budgetKeys, key_description) ?? key_description;
      const dataStore = parsedBankData[amount < 0 ? "spending" : "income"];

      parsedBankData[amount < 0 ? "totalSpending" : "totalIncome"] += amount;

      const data = {
        date: `${date.slice(0, 4)}/${date.slice(4, 6)}/${date.slice(6)}`,
        amount,
      };

      dataStore[key] = {
        totalAmount:
          key in dataStore ? dataStore[key].totalAmount + amount : amount,
        data: key in dataStore ? [...dataStore[key].data, data] : [data],
      };
    });
  }

  return parsedBankData;
};

const consolidateData = (
  budget: Budget,
  bankData: BankData
): {
  expenseSummaryData: ExpenseSummaryData[];
  expandedExpenseData: { key: string; data: AmountData[] }[];
} => {
  const budgetKeys = Object.keys(budget);
  const bankDataSpendingKeys = Object.keys(bankData.spending);
  const expenseSummaryData: ExpenseSummaryData[] = [];

  bankDataSpendingKeys.forEach((key) => {
    const inBudgetKey = !!getDataFromArray(budgetKeys, key);

    expenseSummaryData.push({
      name: inBudgetKey ? budget[key].name : key,
      type: inBudgetKey ? budget[key].expenseType : "No Type",
      budget: inBudgetKey ? budget[key].budget : 0,
      cost: bankData.spending[key].totalAmount,
      diff:
        bankData.spending[key].totalAmount +
        (inBudgetKey ? budget[key].budget : 0),
    });
  });

  const expandedExpenseData = bankDataSpendingKeys
    .filter((key) => bankData.spending[key].data.length > 1)
    .sort(
      (a, b) =>
        bankData.spending[b].data.length - bankData.spending[a].data.length
    )
    .map((key) => ({ key, data: bankData.spending[key].data }));

  return { expenseSummaryData, expandedExpenseData };
};

const addExpandedDataToWorksheet = (
  worksheet: Worksheet,
  col: string,
  row: number,
  { key, data }: { key: string; data: AmountData[] }
) => {
  const columns = ["Date", "Cost"];
  const headerCell = worksheet.getCell(`${col}${row - 1}`);
  worksheet.mergeCells(
    `${col}${row - 1}:${String.fromCharCode(col.charCodeAt(0) + 1)}${row - 1}`
  );
  headerCell.value = key;
  headerCell.style = HEADER_STYLE;

  worksheet.addTable({
    columns: columns.map<TableColumnProperties>((name) => ({
      name,
    })),
    name: key.replace(/[^a-zA-Z]/g, "_").replace(/^_$/g, ""),
    ref: `${col}${row}`,
    rows: data.map((data) => [data.date, data.amount * -1]),
  });
};

const exportData = async (
  expenseSummaryData: ExpenseSummaryData[],
  expandedExpenseData: { key: string; data: AmountData[] }[],
  {
    incomeData,
    totalIncome,
    totalExpense,
  }: {
    incomeData: BankData["income"];
    totalIncome: number;
    totalExpense: number;
  }
) => {
  const workbook = new Workbook();

  await workbook.xlsx.readFile(BUDGET_FILE);
  const intMonth = parseInt(expandedExpenseData[0].data[0].date.split("/")[1]);
  const month = MONTHS[intMonth - 1];

  workbook.removeWorksheet(month);
  const worksheet = await workbook.addWorksheet(month);

  // entire page styling
  Array.from({ length: 20 }).forEach((_, index) => {
    worksheet.getColumn(index + 1).font = { name: "Arial Narrow" };
  });

  //  Insight Summary Section
  const savings = totalIncome + totalExpense; // totalExpense is negative
  worksheet.addRows([
    ["Total Income", totalIncome],
    ["Total Expense", totalExpense * -1],
    ["Savings", savings],
  ]);

  const borderStyle: Partial<Border> = {
    color: BLACK,
    style: "thin",
  };

  worksheet.getCell("B3").style = {
    border: {
      bottom: borderStyle,
      top: borderStyle,
      left: borderStyle,
      right: borderStyle,
    },
    font: {
      name: "Arial Narrow",
    },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: savings > 0 ? GREEN : WHITE,
    },
  };

  // Expense Summary Header
  worksheet.addRows([[], ["Expense Summary"]]);
  worksheet.mergeCells("A5:E5");
  worksheet.getCell("A5").style = MAIN_HEADER_STYLE;

  // Expense Summary Table
  const columnNamesExpenseSummary = [
    "Expense Name",
    "Expense Type",
    "Expense",
    "Budget",
    "Difference",
  ];
  worksheet.addTable({
    columns: columnNamesExpenseSummary.map<TableColumnProperties>(
      (name, index) => ({
        name,
        filterButton: index > 0,
      })
    ),
    name: "Expense_Summary",
    ref: "A6",
    rows: expenseSummaryData.map((expenseSummary) => [
      expenseSummary.name.trim(),
      expenseSummary.type,
      expenseSummary.cost * -1,
      expenseSummary.budget,
      expenseSummary.diff,
    ]),
  });

  // Income Table Header
  worksheet.addRows([[], ["Income Sources"]]);
  let currentRow = expenseSummaryData.length + 8;
  worksheet.mergeCells(`A${currentRow}:C${currentRow}`);
  worksheet.getCell(`A${currentRow}`).style = MAIN_HEADER_STYLE;

  // Income Table
  const columnNamesIncome = ["Date", "From", "Amount"];
  worksheet.addTable({
    columns: columnNamesIncome.map<TableColumnProperties>((name) => ({ name })),
    name: "Income",
    ref: `A${9 + expenseSummaryData.length}`,
    rows: Object.keys(incomeData)
      .map((key) =>
        incomeData[key].data.map((data) => [data.date, key.trim(), data.amount])
      )
      .reduce((flatten, arr) => [...flatten, ...arr]),
  });

  // Expense Break Down Header
  worksheet.getCell("J1").value = "Expense Break Down";
  worksheet.mergeCells("J1:S1");
  worksheet.getCell("J1").style = MAIN_HEADER_STYLE;

  let rowToPasteData = [3, 3, 3];
  expandedExpenseData.forEach((data, index) => {
    const columns = ["J", "N", "R"];
    const indx = index % 3;
    addExpandedDataToWorksheet(
      worksheet,
      columns[indx],
      rowToPasteData[indx],
      data
    );
    rowToPasteData[indx] += data.data.length + 3;
  });

  // Resize columns and format all numbers to amount
  worksheet.columns.forEach((column) => {
    const lengths: number[] = [];
    if (column.eachCell) {
      column.eachCell((cell) => {
        if (cell.value) {
          const cellValue = cell.value.toString();
          lengths.push(cellValue.length);

          if (cellValue.match(/^-?\d+(\.\d+)?$/)) {
            cell.numFmt = "$#,##0.00;[Red]-$#,##0.00";
          }
        }
      });
    }

    if (lengths.length > 0) {
      column.width = lengths.reduce((p, c) => (p > c ? p : c));
    }
  });

  // Write to and Open Workbook
  await workbook.xlsx.writeFile(BUDGET_FILE);
  open(BUDGET_FILE);
};

const main = async (filePath: string) => {
  const budget = await getBudget();

  if (!budget) {
    console.log("No Budget data was found. Check location of template file.");
    return;
  }

  const bankData = await getBankData(filePath, Object.keys(budget));

  if (
    !bankData ||
    (Object.keys(bankData.spending).length === 0 &&
      Object.keys(bankData.income).length === 0)
  ) {
    console.log(
      "Either no file was found in the location or file path is wrong."
    );
    return;
  }

  const { expenseSummaryData, expandedExpenseData } = consolidateData(
    budget,
    bankData
  );

  exportData(expenseSummaryData, expandedExpenseData, {
    incomeData: bankData.income,
    totalIncome: bankData.totalIncome,
    totalExpense: bankData.totalSpending,
  });
};

main(`./Data`);
