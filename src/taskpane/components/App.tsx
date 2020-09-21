import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

import { IStackTokens, Stack } from "office-ui-fabric-react/lib/Stack";
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { getWorksheetNames } from "../utils/index";

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 }
};

const options: IDropdownOption[] = [
  { key: "sheets", text: "Sheets", itemType: DropdownMenuItemType.Header },
  { key: "CompaniesCurrent", text: "CompaniesCurrent" },
  { key: "BadgerCompanyAccounts", text: "BadgerCompanyAccounts" },
  { key: "BadgerCheckins", text: "BadgerCheckins" },
  { key: "BadgerContacts", text: "BadgerContacts" }
];

const stackTokens: IStackTokens = { childrenGap: 20 };

/*
  React.useEffect(() => {
    async function getWorksheetDataCopy(sheetToImport) {
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(sheetToImport);
        var largeRange = context.workbook.getSelectedRange();
        largeRange.load(["rowCount", "columnCount"]);
        await context.sync();

        const range = sheet.getRange(`A2:A${largeRange.rowCount}`);
        range.load("values");
        await context.sync();

        console.log("range", JSON.stringify(range.values, null, 4));
        //console.log("Fuse", Fuse);
        //const sheetCopy = context.workbook.worksheets.add("badger-company49759-copy");

        //queueCommandsToCreateTemperatureTable(sheet);
        //sheet.activate();

        await context.sync();
        //console.log("range", JSON.stringify(range.text, null, 4));
      });
    }

    getWorksheetDataCopy(sheetToImport);
  }, [sheetToImport]);*/

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

const App: React.FC<AppProps> = ({ title, isOfficeInitialized }) => {
  const [items, setListItems] = React.useState({ listItems: [] });
  const [sheetToImport, setSheetToImport] = React.useState({ key: "", text: "" });
  const [sheetVals, setSheetVals] = React.useState("Nothing Yet");
  const [error, setError] = React.useState("No current errors");
  const onChange = (event, item) => {
    console.log("Event item", event, item);
    setSheetToImport({ key: "BadgerContacts", text: "BadgerContacts" });
  };

  React.useEffect(() => {
    setListItems({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Advanced Duplication Detection"
        },
        {
          icon: "Unlock",
          primaryText: "Address and Name parser"
        },
        {
          icon: "Design",
          primaryText: ""
        }
      ]
    });
  }, [items]);

  const click = async () => {
    try {
      //const tempSheetVals = await getWorksheetData(sheetToImport.text);
      const worksheetNames = await getWorksheetNames();
      setSheetVals(JSON.stringify(worksheetNames, null, 4));
      /*await Excel.run(async context => {
       
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });*/
    } catch (error) {
      setSheetVals(JSON.stringify(error, null, 4));
      setError(JSON.stringify(error, null, 4));
      console.error(error);
    }
  };

  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/repfabric-logo.png" message="Please sideload your addin to see app body." />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo="assets/repfabric-logo.png" title={title} message="Repfabric" />
      <HeroList message="Repfabric data conversion tools" items={items.listItems}>
        <Stack tokens={stackTokens}>
          <Dropdown
            placeholder="Select a sheet"
            label="Sheet To Import"
            selectedKey={sheetToImport?.key || undefined}
            onChange={onChange}
            options={options}
            styles={dropdownStyles}
          />
        </Stack>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={click}
        >
          Run {sheetToImport?.text ?? "No sheet selected"}
        </Button>
        <p className="ms-font-l">Error</p>
        <code>{error}</code>
        <p className="ms-font-l">SheetVals</p>
        <p className="ms-font-l">{sheetVals}</p>
      </HeroList>
    </div>
  );
};

export default App;
