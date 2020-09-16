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

  const onChange = event => {
    setSheetToImport(event.target.value);
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

  const getWorksheetDataCopy = async sheetToImport => {
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
  };

  const click = async () => {
    try {
      await getWorksheetDataCopy(sheetToImport);
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo="assets/logo-filled.png" title={title} message="Repfabric" />
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
          Run
        </Button>
      </HeroList>
    </div>
  );
};

export default App;
