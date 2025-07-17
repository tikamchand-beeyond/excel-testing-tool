import React from "react";
import { PrimaryButton, Stack } from "@fluentui/react";

const TaskPane: React.FC = () => {
  const [log, setLog] = React.useState<string>("");

  /* Helper that wraps every Excel call in Excel.run */
  const run = async (work: (ctx: Excel.RequestContext) => Promise<string>) => {
    try {
      const message = await Excel.run(work);
      setLog(message);
    } catch (error) {
      setLog(`Error: ${error}`);
    }
  };

  /* 1. Read A1 */
  const readA1 = () =>
    run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.load("values");
      await context.sync();
      const value = range.values[0][0];
      return `A1 = ${value ?? "<empty>"}`;
    });

  /* 2. Write timestamp to A2 */
  const writeTimestamp = () =>
    run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getRange("A2").values = [[new Date().toLocaleString()]];
      return "Timestamp written to A2";
    });

  /* 3. Create / extend a sample table */
  const createTable = () =>
    run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Create if it doesn’t exist
      let table = sheet.tables.getItemOrNullObject("TestTable");
      await context.sync();
      if (!table.isNullObject) {
        table.rows.add(null, [["Charlie", 95]]);
        return "Row added to existing TestTable";
      }

      // Build initial table
      const range = sheet.getRange("B1:C3");
      range.values = [["Name", "Score"], ["Alice", 90], ["Bob", 85]];
      const newTable = sheet.tables.add(range, true);
      newTable.name = "TestTable";
      newTable.rows.add(null, [["Charlie", 95]]);
      return "Table 'TestTable' created and row added";
    });

  return (
    <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20 }}>
      <h2>Excel Add-in Test Tool</h2>

      <PrimaryButton onClick={readA1} text="Read A1" />
      <PrimaryButton onClick={writeTimestamp} text="Write timestamp to A2" />
      <PrimaryButton onClick={createTable} text="Create / extend table" />
      <PrimaryButton onClick={() => setLog("")} text="Clear log" />

      <div style={{ fontSize: 12, whiteSpace: "pre-wrap" }}>{log}</div>
    </Stack>
  );
};

export default TaskPane;