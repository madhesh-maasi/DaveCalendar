import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Context } from "react";

export interface ICalendarProps {
  description: string;
  context: WebPartContext;
  graphcontext: any;
  spcontext: any;
}
