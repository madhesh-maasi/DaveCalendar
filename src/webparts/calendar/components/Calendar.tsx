import * as React from "react";
import styles from "./Calendar.module.scss";
import { ICalendarProps } from "./ICalendarProps";
import CalendarApp from "./CalendarApp";
import "./Calendar.css";
export default class Calendar extends React.Component<ICalendarProps, {}> {
  public render(): React.ReactElement<ICalendarProps> {
    return (
      <div className="calendar-section">
        <CalendarApp
          spcontext={this.props.spcontext}
          graphcontext={this.props.graphcontext}
        />
      </div>
    );
  }
}
