import * as React from "react";
import styles from "./Calendar.module.scss";
import { ICalendarProps } from "./ICalendarProps";
import { escape } from "@microsoft/sp-lodash-subset";
import CalendarColorView from "./CalendarColorView";
import CalendarDetails from "./CalendarDetails";
import { MSGraphClient } from "@microsoft/sp-http";
import "./Calendar.css";
let arrColorDetails = [];
export default class Calendar extends React.Component<ICalendarProps, {}> {
  public render(): React.ReactElement<ICalendarProps> {
    // thisContext = this.context;
    this.props.spcontext.web.lists
      .getByTitle("CalColorConfig")
      .items.get()
      .then((liData) => {
        arrColorDetails = liData;
      });
    let classes = styles.calendar;
    return (
      <div className="calendar-section">
        <CalendarColorView spcontext={this.props.spcontext}></CalendarColorView>
        <CalendarDetails
          spcontext={this.props.spcontext}
          graphcontext={this.props.graphcontext}
        />
      </div>
    );
  }
}
