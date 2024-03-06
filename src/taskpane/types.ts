import { FC } from "react";

export interface AppProps {
  title: string;
}

export interface AppTabProps {
  title: string;
  active: boolean;
}

export interface AppTab {
  value: string;
  label: string;
  TabComponent: FC<AppTabProps>;
}
