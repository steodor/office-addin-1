import React, { useState } from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import {
  SelectTabData,
  SelectTabEvent,
  SelectTabEventHandler,
  Tab,
  TabList,
  TabValue,
  makeStyles,
  shorthands,
} from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import useEvent from "react-use-event-hook";
import classNames from "classnames";
import ReferenceInsertion from "./ReferenceInsertion";
import { AppProps, AppTab, AppTabProps } from "../types";

const appTabs: AppTab[] = [
  {
    value: "tab1",
    label: "Ddroidd",
    TabComponent: (props: AppTabProps) => {
      // The list items are static and won't change at runtime,
      //   so this should be an ordinary const, not a part of state.
      const listItems: HeroListItem[] = [
        {
          icon: <Ribbon24Regular />,
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: <LockOpen24Regular />,
          primaryText: "Unlock features and functionality",
        },
        {
          icon: <DesignIdeas24Regular />,
          primaryText: "Create and visualize like a pro",
        },
      ];

      return (
        <>
          <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
          <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
          <TextInsertion />
        </>
      );
    },
  },
  {
    value: "tab2",
    label: "Reference",
    TabComponent: (props: AppTabProps) => {
      return (
        <>
          <ReferenceInsertion {...props} />
        </>
      );
    },
  },
];

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  tabContent: {
    display: "none",
    ...shorthands.padding("16px"),
  },
  active: {
    display: "block",
  },
});

const App = (props: AppProps) => {
  const styles = useStyles();

  const [activeTab, setActiveTab] = useState<TabValue>(appTabs[0].value);

  const onTabSelect = useEvent<SelectTabEventHandler>((_event: SelectTabEvent<HTMLElement>, data: SelectTabData) => {
    setActiveTab(data.value);
  });

  return (
    <div className={styles.root}>
      <TabList selectedValue={activeTab} onTabSelect={onTabSelect}>
        {appTabs.map((tab: AppTab) => (
          <Tab key={tab.value} value={tab.value}>
            {tab.label}
          </Tab>
        ))}
      </TabList>
      {appTabs.map((tab: AppTab) => (
        <div
          key={tab.value}
          className={classNames(styles.tabContent, {
            [styles.active]: activeTab == tab.value,
          })}
        >
          <tab.TabComponent {...props} active={activeTab == tab.value} />
        </div>
      ))}
    </div>
  );
};

export default App;
