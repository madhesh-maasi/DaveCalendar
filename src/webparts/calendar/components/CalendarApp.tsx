import * as React from "react";
import { useState, useEffect } from "react";
import CalendarColorView from "./CalendarColorView";
import CalendarDetails from "./CalendarDetails";
import { graph } from "@pnp/graph";
import xmlToJSON from "xmltojson";
import { loadTheme } from "office-ui-fabric-react";

// var convert = require("xml-js");

let data = [];
let allData = [];
let FilteredData = [];
let arrColorVar = [];
let userInGroup = false;
// let timeZone = "Pacific Standard Time"; // For Dave RPM
// let timeZone = "Eastern Standard Time"; // For SilverLeaf and EJF
let timeZone = "India Standard Time"; //for local time zone
let headers = { Prefer: 'outlook.timezone="' + timeZone + '"' };
let isOnload = true;
const CalendarApp = (props) => {
  const [arrColor, setArrColor] = useState([]);
  const [events, setEvents] = useState([]);
  const [selectedEvent, setSelectedEvent] = useState([]);
  //   useState
  // colorList calls
  // ! All Calendar Event Function
  const getEventsFromCalendar = () => {
    props.spcontext.web.lists
      .getByTitle("CalColorConfig")
      .items.get()
      .then((colorLi) => {
        arrColorVar = colorLi;
        setArrColor(arrColorVar);
      })
      .then(() => {
        props.spcontext.web.lists
          .getByTitle("CalMonthConfig")
          .items.get()
          .then(async (li) => {
            let date = new Date();
            let firstDay: any = new Date(
              date.getFullYear(),
              date.getMonth(),
              1
            );
            let lastDay: any = new Date(
              date.getFullYear(),
              date.getMonth() + 1,
              0
            );

            li.length > 0 && li[0].Month && li[0].Month != null
              ? firstDay.setMonth(firstDay.getMonth() - li[0].Month)
              : firstDay.setMonth(firstDay.getMonth() - 0);
            li.length > 0 && li[0].Month && li[0].Month != null
              ? lastDay.setMonth(lastDay.getMonth() + li[0].Month)
              : lastDay.setMonth(lastDay.getMonth() + 0);

            let firstDayOfMonth =
              new Date(firstDay).toISOString().split("T")[0] + "T12:00:00.000Z";
            let LastDayOfMonth =
              new Date(lastDay).toISOString().split("T")[0] + "T12:00:00.000Z";
            let myId = "";
            let currEmail = "";
            await graph.me().then((myR: any) => {
              myId = myR.id;
              currEmail = myR.userPrincipalName;
            });

            await graph.me.events
              .configure({ headers })
              .filter(
                "start/datetime ge '" +
                  firstDayOfMonth +
                  "' and end/datetime le '" +
                  LastDayOfMonth +
                  "'"
              )
              .top(999)()
              .then((event) => {
                console.log(event);

                data = event.map((evt) => {
                  let recED;
                  let recEDate;
                  let recEndDateTime;
                  let myEventColor = arrColorVar.filter(
                    (aC) => aC.IsUser == true
                  )[0].HexCode;
                  let dow = [];
                  evt.recurrence &&
                  evt.recurrence.pattern.type == "weekly" &&
                  evt.recurrence.pattern.daysOfWeek.length > 0
                    ? evt.recurrence.pattern.daysOfWeek.forEach((dw) => {
                        dw == "monday"
                          ? dow.push(1)
                          : dw == "tuesday"
                          ? dow.push(2)
                          : dw == "wednesday"
                          ? dow.push(3)
                          : dw == "thursday"
                          ? dow.push(4)
                          : dw == "friday"
                          ? dow.push(5)
                          : dw == "saturday"
                          ? dow.push(6)
                          : dw == "sunday"
                          ? dow.push(7)
                          : "";
                        recEDate = new Date(evt.recurrence.range.endDate);
                        recEDate.setDate(recEDate.getDate() + 1);
                        recED = recEDate.toISOString().split("T")[0];
                        recEndDateTime = `${recED}T${
                          evt.end.dateTime.split("T")[1]
                        }`;
                      })
                    : evt.recurrence && evt.recurrence.pattern.type == "daily"
                    ? ((recEDate = new Date(evt.recurrence.range.endDate)),
                      (recED = recEDate.toISOString().split("T")[0]),
                      (recEndDateTime = `${recED}T${
                        evt.end.dateTime.split("T")[1]
                      }`))
                    : "";

                  return evt.recurrence &&
                    evt.recurrence.pattern.type == "weekly"
                    ? {
                        id: evt.id,
                        // daysOfWeek: dow,
                        // startRecur: evt.recurrence.range.startDate,
                        // endRecur: recED,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
                        rrule: {
                          freq: "weekly",
                          interval: evt.recurrence.pattern.interval,
                          byweekday: dow.map((dw) =>
                            dw == 1
                              ? "mo"
                              : dw == 2
                              ? "tu"
                              : dw == 3
                              ? "we"
                              : dw == 4
                              ? "th"
                              : dw == 5
                              ? "fr"
                              : dw == 6
                              ? "sa"
                              : "su"
                          ),
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : evt.recurrence.range.endDate, // will also accept '20120201'
                        },
                      }
                    : evt.recurrence && evt.recurrence.pattern.type == "daily"
                    ? {
                        id: evt.id,
                        // daysOfWeek: [1, 2, 3, 4, 5, 6, 7],
                        // startRecur: evt.recurrence.range.startDate,
                        // endRecur: recED,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        rrule: {
                          freq: "daily",
                          interval: evt.recurrence.pattern.interval,
                          byweekday: dow.map((dw) =>
                            dw == 1
                              ? "mo"
                              : dw == 2
                              ? "tu"
                              : dw == 3
                              ? "we"
                              : dw == 4
                              ? "th"
                              : dw == 5
                              ? "fr"
                              : dw == 6
                              ? "sa"
                              : "su"
                          ),
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : `${evt.recurrence.range.endDate}T${
                                  evt.end.dateTime.split("T")[1]
                                }`, // will also accept '20120201'
                        },
                        itemFrom: "PersonalCalendar",
                      }
                    : evt.recurrence &&
                      evt.recurrence.pattern.type == "absoluteMonthly"
                    ? {
                        id: evt.id,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        rrule: {
                          freq: "monthly",
                          interval: evt.recurrence.pattern.interval,
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : evt.recurrence.range.endDate, // will also accept '20120201'
                        },
                        itemFrom: "PersonalCalendar",
                      }
                    : evt.recurrence &&
                      evt.recurrence.pattern.type == "relativeMonthly"
                    ? {
                        id: evt.id,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
                        rrule: {
                          freq: "monthly",
                          interval: evt.recurrence.pattern.interval,
                          // index: evt.recurrence.pattern.index,
                          byweekday: evt.recurrence.pattern.daysOfWeek.map(
                            (day) =>
                              day == "monday"
                                ? "mo"
                                : day == "tuesday"
                                ? "tu"
                                : day == "wednesday"
                                ? "we"
                                : day == "thursday"
                                ? "th"
                                : day == "friday"
                                ? "fr"
                                : day == "saturday"
                                ? "sa"
                                : day == "sunday"
                                ? "su"
                                : ""
                          ),
                          bysetpos:
                            evt.recurrence.pattern.index == "first"
                              ? 1
                              : evt.recurrence.pattern.index == "second"
                              ? 2
                              : evt.recurrence.pattern.index == "third"
                              ? 3
                              : evt.recurrence.pattern.index == "fourth"
                              ? 4
                              : -1,
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : evt.recurrence.range.endDate, // will also accept '20120201'
                        },
                        //  description: evt.bodyPreview,
                      }
                    : {
                        id: evt.id,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end: evt.end.dateTime,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
                      };
                });
                console.log(data);
              })
              .then(async () => {
                li.length > 0 && li[0].GroupID != null
                  ? await graph.groups
                      .getById(li[0].GroupID)
                      .members.top(999)()
                      .then(async (groupRes: any) => {
                        userInGroup =
                          groupRes.filter((gR) => gR.id == myId).length > 0;
                        userInGroup
                          ? await graph.groups
                              .getById(li[0].GroupID)
                              .events.configure({ headers })
                              .filter(
                                "start/datetime ge '" +
                                  firstDayOfMonth +
                                  "' and end/datetime le '" +
                                  LastDayOfMonth +
                                  "'"
                              )
                              .top(999)()
                              .then((result: any) => {
                                console.log(result);
                                let data1 = [];
                                let recEndDateTime;
                                let recEDate;
                                data1 = result.map((evt) => {
                                  let recED = "";
                                  let eventColor = "";
                                  let eventType = "";
                                  let eventColorArr = arrColorVar.filter(
                                    (colLi) => {
                                      return evt.subject
                                        .toLowerCase()
                                        .includes(colLi.Title.toLowerCase());
                                    }
                                  );
                                  eventColorArr.length > 0
                                    ? ((eventColor = eventColorArr[0].HexCode),
                                      (eventType = "GroupCalendar"))
                                    : ((eventColor = arrColorVar.filter(
                                        (colLi) =>
                                          colLi.DefaultEventColor == true
                                      )[0].HexCode),
                                      (eventType = "GroupCalendar Other"));

                                  let dow = [];
                                  evt.recurrence &&
                                  evt.recurrence.pattern.type == "weekly" &&
                                  evt.recurrence.pattern.daysOfWeek.length > 0
                                    ? evt.recurrence.pattern.daysOfWeek.forEach(
                                        (dw) => {
                                          dw == "monday"
                                            ? dow.push(1)
                                            : dw == "tuesday"
                                            ? dow.push(2)
                                            : dw == "wednesday"
                                            ? dow.push(3)
                                            : dw == "thursday"
                                            ? dow.push(4)
                                            : dw == "friday"
                                            ? dow.push(5)
                                            : dw == "saturday"
                                            ? dow.push(6)
                                            : dw == "sunday"
                                            ? dow.push(7)
                                            : "";
                                          recEDate = new Date(
                                            evt.recurrence.range.endDate
                                          );
                                          recEDate.setDate(
                                            recEDate.getDate() + 1
                                          );
                                          recED = recEDate
                                            .toISOString()
                                            .split("T")[0];
                                          recEDate = new Date(
                                            evt.recurrence.range.endDate
                                          );
                                          recEDate.setDate(
                                            recEDate.getDate() + 1
                                          );
                                          recED = recEDate
                                            .toISOString()
                                            .split("T")[0];
                                          recEndDateTime = `${recED}T${
                                            evt.end.dateTime.split("T")[1]
                                          }`;
                                        }
                                      )
                                    : evt.recurrence &&
                                      evt.recurrence.pattern.type == "daily"
                                    ? ((recEDate = new Date(
                                        evt.recurrence.range.endDate
                                      )),
                                      (recED = recEDate
                                        .toISOString()
                                        .split("T")[0]),
                                      (recEndDateTime = `${recED}T${
                                        evt.end.dateTime.split("T")[1]
                                      }`))
                                    : "";
                                  return evt.recurrence &&
                                    evt.recurrence.pattern.type == "weekly"
                                    ? {
                                        id: evt.id,
                                        title: evt.subject,
                                        start: evt.start.dateTime,
                                        end:
                                          evt.recurrence.range.type ==
                                            "noEnd" ||
                                          evt.recurrence.range.endDate ==
                                            "0001-01-01"
                                            ? `${
                                                new Date(
                                                  new Date(
                                                    evt.recurrence.range.startDate
                                                  ).setFullYear(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).getFullYear() + 1
                                                  )
                                                )
                                                  .toISOString()
                                                  .split("T")[0]
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`
                                            : `${
                                                evt.recurrence.range.endDate
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`,
                                        initialDate: evt.start.dateTime,
                                        display: "block",
                                        attendees: evt.attendees,
                                        description: evt.bodyPreview,
                                        backgroundColor: eventColor,
                                        borderColor: eventColor,
                                        allDay: evt.isAllDay,
                                        itemFrom: eventType,
                                        rrule: {
                                          freq: "weekly",
                                          interval:
                                            evt.recurrence.pattern.interval,
                                          byweekday: dow.map((dw) =>
                                            dw == 1
                                              ? "mo"
                                              : dw == 2
                                              ? "tu"
                                              : dw == 3
                                              ? "we"
                                              : dw == 4
                                              ? "th"
                                              : dw == 5
                                              ? "fr"
                                              : dw == 6
                                              ? "sa"
                                              : "su"
                                          ),
                                          dtstart: `${
                                            evt.recurrence.range.startDate
                                          }T${
                                            evt.start.dateTime.split("T")[1]
                                          }`, // will also accept '20120201T103000'
                                          until:
                                            evt.recurrence.range.type ==
                                              "noEnd" ||
                                            evt.recurrence.range.endDate ==
                                              "0001-01-01"
                                              ? `${
                                                  new Date(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).setFullYear(
                                                      new Date(
                                                        evt.recurrence.range.startDate
                                                      ).getFullYear() + 1
                                                    )
                                                  )
                                                    .toISOString()
                                                    .split("T")[0]
                                                }T${
                                                  evt.end.dateTime.split("T")[1]
                                                }`
                                              : evt.recurrence.range.endDate, // will also accept '20120201'
                                        },
                                      }
                                    : evt.recurrence &&
                                      evt.recurrence.pattern.type == "daily"
                                    ? {
                                        id: evt.id,
                                        title: evt.subject,
                                        start: evt.start.dateTime,
                                        end:
                                          evt.recurrence.range.type ==
                                            "noEnd" ||
                                          evt.recurrence.range.endDate ==
                                            "0001-01-01"
                                            ? `${
                                                new Date(
                                                  new Date(
                                                    evt.recurrence.range.startDate
                                                  ).setFullYear(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).getFullYear() + 1
                                                  )
                                                )
                                                  .toISOString()
                                                  .split("T")[0]
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`
                                            : `${
                                                evt.recurrence.range.endDate
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`,
                                        // end: recEndDateTime,
                                        display: "block",
                                        attendees: evt.attendees,
                                        description: evt.bodyPreview,
                                        backgroundColor: eventColor,
                                        borderColor: eventColor,
                                        allDay: evt.isAllDay,
                                        itemFrom: eventType,
                                        rrule: {
                                          freq: "daily",
                                          interval:
                                            evt.recurrence.pattern.interval,
                                          byweekday: dow.map((dw) =>
                                            dw == 1
                                              ? "mo"
                                              : dw == 2
                                              ? "tu"
                                              : dw == 3
                                              ? "we"
                                              : dw == 4
                                              ? "th"
                                              : dw == 5
                                              ? "fr"
                                              : dw == 6
                                              ? "sa"
                                              : "su"
                                          ),
                                          dtstart: `${
                                            evt.recurrence.range.startDate
                                          }T${
                                            evt.start.dateTime.split("T")[1]
                                          }`, // will also accept '20120201T103000'
                                          until:
                                            evt.recurrence.range.type ==
                                              "noEnd" ||
                                            evt.recurrence.range.endDate ==
                                              "0001-01-01"
                                              ? `${
                                                  new Date(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).setFullYear(
                                                      new Date(
                                                        evt.recurrence.range.startDate
                                                      ).getFullYear() + 1
                                                    )
                                                  )
                                                    .toISOString()
                                                    .split("T")[0]
                                                }T${
                                                  evt.end.dateTime.split("T")[1]
                                                }`
                                              : `${
                                                  evt.recurrence.range.endDate
                                                }T${
                                                  evt.end.dateTime.split("T")[1]
                                                }`, // will also accept '20120201'
                                        },
                                        //  description: evt.bodyPreview,
                                      }
                                    : evt.recurrence &&
                                      evt.recurrence.pattern.type ==
                                        "absoluteMonthly"
                                    ? {
                                        id: evt.id,
                                        title: evt.subject,
                                        // daysOfWeek: [1, 2, 3, 4, 5, 6, 7],
                                        dayOfMonth:
                                          evt.recurrence.pattern.dayOfMonth,

                                        start: evt.start.dateTime,
                                        end:
                                          evt.recurrence.range.type ==
                                            "noEnd" ||
                                          evt.recurrence.range.endDate ==
                                            "0001-01-01"
                                            ? `${
                                                new Date(
                                                  new Date(
                                                    evt.recurrence.range.startDate
                                                  ).setFullYear(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).getFullYear() + 1
                                                  )
                                                )
                                                  .toISOString()
                                                  .split("T")[0]
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`
                                            : `${
                                                evt.recurrence.range.endDate
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`,
                                        // end: evt.end.dateTime,
                                        display: "block",
                                        attendees: evt.attendees,
                                        description: evt.bodyPreview,
                                        backgroundColor: eventColor,
                                        borderColor: eventColor,
                                        allDay: evt.isAllDay,
                                        itemFrom: eventType,
                                        rrule: {
                                          freq: "monthly",
                                          interval:
                                            evt.recurrence.pattern.interval,
                                          dtstart: `${
                                            evt.recurrence.range.startDate
                                          }T${
                                            evt.start.dateTime.split("T")[1]
                                          }`, // will also accept '20120201T103000'
                                          until:
                                            evt.recurrence.range.type ==
                                              "noEnd" ||
                                            evt.recurrence.range.endDate ==
                                              "0001-01-01"
                                              ? `${
                                                  new Date(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).setFullYear(
                                                      new Date(
                                                        evt.recurrence.range.startDate
                                                      ).getFullYear() + 1
                                                    )
                                                  )
                                                    .toISOString()
                                                    .split("T")[0]
                                                }T${
                                                  evt.end.dateTime.split("T")[1]
                                                }`
                                              : evt.recurrence.range.endDate, // will also accept '20120201'
                                        },
                                        //  description: evt.bodyPreview,
                                      }
                                    : evt.recurrence &&
                                      evt.recurrence.pattern.type ==
                                        "relativeMonthly"
                                    ? {
                                        id: evt.id,
                                        title: evt.subject,
                                        start: evt.start.dateTime,
                                        end:
                                          evt.recurrence.range.type ==
                                            "noEnd" ||
                                          evt.recurrence.range.endDate ==
                                            "0001-01-01"
                                            ? `${
                                                new Date(
                                                  new Date(
                                                    evt.recurrence.range.startDate
                                                  ).setFullYear(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).getFullYear() + 1
                                                  )
                                                )
                                                  .toISOString()
                                                  .split("T")[0]
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`
                                            : `${
                                                evt.recurrence.range.endDate
                                              }T${
                                                evt.end.dateTime.split("T")[1]
                                              }`, // will also accept '20120201',
                                        // end: evt.end.dateTime,
                                        display: "block",
                                        attendees: evt.attendees,
                                        description: evt.bodyPreview,
                                        backgroundColor: eventColor,
                                        borderColor: eventColor,
                                        allDay: evt.isAllDay,
                                        itemFrom: eventType,
                                        rrule: {
                                          freq: "monthly",
                                          interval:
                                            evt.recurrence.pattern.interval,
                                          // index: evt.recurrence.pattern.index,
                                          byweekday:
                                            evt.recurrence.pattern.daysOfWeek.map(
                                              (day) =>
                                                day == "monday"
                                                  ? "mo"
                                                  : day == "tuesday"
                                                  ? "tu"
                                                  : day == "wednesday"
                                                  ? "we"
                                                  : day == "thursday"
                                                  ? "th"
                                                  : day == "friday"
                                                  ? "fr"
                                                  : day == "saturday"
                                                  ? "sa"
                                                  : day == "sunday"
                                                  ? "su"
                                                  : ""
                                            ),
                                          bysetpos:
                                            evt.recurrence.pattern.index ==
                                            "first"
                                              ? 1
                                              : evt.recurrence.pattern.index ==
                                                "second"
                                              ? 2
                                              : evt.recurrence.pattern.index ==
                                                "third"
                                              ? 3
                                              : evt.recurrence.pattern.index ==
                                                "fourth"
                                              ? 4
                                              : -1,
                                          dtstart: `${
                                            evt.recurrence.range.startDate
                                          }T${
                                            evt.start.dateTime.split("T")[1]
                                          }`, // will also accept '20120201T103000'
                                          until:
                                            evt.recurrence.range.type ==
                                              "noEnd" ||
                                            evt.recurrence.range.endDate ==
                                              "0001-01-01"
                                              ? `${
                                                  new Date(
                                                    new Date(
                                                      evt.recurrence.range.startDate
                                                    ).setFullYear(
                                                      new Date(
                                                        evt.recurrence.range.startDate
                                                      ).getFullYear() + 1
                                                    )
                                                  )
                                                    .toISOString()
                                                    .split("T")[0]
                                                }T${
                                                  evt.end.dateTime.split("T")[1]
                                                }`
                                              : evt.recurrence.range.endDate, // will also accept '20120201'
                                        },
                                        //  description: evt.bodyPreview,
                                      }
                                    : {
                                        id: evt.id,
                                        title: evt.subject,
                                        start: evt.start.dateTime,
                                        end: evt.end.dateTime,
                                        display: "block",
                                        attendees: evt.attendees,
                                        description: evt.bodyPreview,
                                        backgroundColor: eventColor,
                                        borderColor: eventColor,
                                        allDay: evt.isAllDay,
                                        itemFrom: eventType,
                                        //  description: evt.bodyPreview,
                                      };
                                });
                                console.log(data1);
                                data = [...data, ...data1];
                                setEvents(data);
                              })
                          : "";
                      })
                  : setEvents(data);
                console.log(data);
              });
          });
      });
  };
  const getEventsFromEvents = () => {
    props.spcontext.web.lists
      .getByTitle("CalColorConfig")
      .items.get()
      .then((colorLi) => {
        arrColorVar = colorLi;
        setArrColor(arrColorVar);
      })
      .then(() => {
        props.spcontext.web.lists
          .getByTitle("CalMonthConfig")
          .items.get()
          .then(async (li) => {
            let date = new Date();
            let firstDay: any = new Date(
              date.getFullYear(),
              date.getMonth(),
              1
            );
            let lastDay: any = new Date(
              date.getFullYear(),
              date.getMonth() + 1,
              0
            );

            li.length > 0 && li[0].Month && li[0].Month != null
              ? firstDay.setMonth(firstDay.getMonth() - li[0].Month)
              : firstDay.setMonth(firstDay.getMonth() - 0);
            li.length > 0 && li[0].Month && li[0].Month != null
              ? lastDay.setMonth(lastDay.getMonth() + li[0].Month)
              : lastDay.setMonth(lastDay.getMonth() + 0);

            let firstDayOfMonth =
              new Date(firstDay).toISOString().split("T")[0] + "T12:00:00.000Z";
            let LastDayOfMonth =
              new Date(lastDay).toISOString().split("T")[0] + "T12:00:00.000Z";
            let myId = "";
            let currEmail = "";
            await graph.me().then((myR: any) => {
              myId = myR.id;
              currEmail = myR.userPrincipalName;
            });
            await graph.me.events
              .configure({ headers })
              .filter(
                "start/datetime ge '" +
                  firstDayOfMonth +
                  "' and end/datetime le '" +
                  LastDayOfMonth +
                  "'"
              )
              .top(999)()
              .then((event) => {
                console.log(event);

                data = event.map((evt) => {
                  let recED;
                  let recEDate;
                  let recEndDateTime;
                  let myEventColor = arrColorVar.filter(
                    (aC) => aC.IsUser == true
                  )[0].HexCode;
                  let dow = [];
                  evt.recurrence &&
                  evt.recurrence.pattern.type == "weekly" &&
                  evt.recurrence.pattern.daysOfWeek.length > 0
                    ? evt.recurrence.pattern.daysOfWeek.forEach((dw) => {
                        dw == "monday"
                          ? dow.push(1)
                          : dw == "tuesday"
                          ? dow.push(2)
                          : dw == "wednesday"
                          ? dow.push(3)
                          : dw == "thursday"
                          ? dow.push(4)
                          : dw == "friday"
                          ? dow.push(5)
                          : dw == "saturday"
                          ? dow.push(6)
                          : dw == "sunday"
                          ? dow.push(7)
                          : "";
                        recEDate = new Date(evt.recurrence.range.endDate);
                        recEDate.setDate(recEDate.getDate() + 1);
                        recED = recEDate.toISOString().split("T")[0];
                        recEndDateTime = `${recED}T${
                          evt.end.dateTime.split("T")[1]
                        }`;
                      })
                    : evt.recurrence && evt.recurrence.pattern.type == "daily"
                    ? ((recEDate = new Date(evt.recurrence.range.endDate)),
                      (recED = recEDate.toISOString().split("T")[0]),
                      (recEndDateTime = `${recED}T${
                        evt.end.dateTime.split("T")[1]
                      }`))
                    : "";

                  return evt.recurrence &&
                    evt.recurrence.pattern.type == "weekly"
                    ? {
                        id: evt.id,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
                        rrule: {
                          freq: "weekly",
                          interval: evt.recurrence.pattern.interval,
                          byweekday: dow.map((dw) =>
                            dw == 1
                              ? "mo"
                              : dw == 2
                              ? "tu"
                              : dw == 3
                              ? "we"
                              : dw == 4
                              ? "th"
                              : dw == 5
                              ? "fr"
                              : dw == 6
                              ? "sa"
                              : "su"
                          ),
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : evt.recurrence.range.endDate, // will also accept '20120201'
                        },
                      }
                    : evt.recurrence && evt.recurrence.pattern.type == "daily"
                    ? {
                        id: evt.id,
                        // daysOfWeek: [1, 2, 3, 4, 5, 6, 7],
                        // startRecur: evt.recurrence.range.startDate,
                        // endRecur: recED,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        rrule: {
                          freq: "daily",
                          interval: evt.recurrence.pattern.interval,
                          byweekday: dow.map((dw) =>
                            dw == 1
                              ? "mo"
                              : dw == 2
                              ? "tu"
                              : dw == 3
                              ? "we"
                              : dw == 4
                              ? "th"
                              : dw == 5
                              ? "fr"
                              : dw == 6
                              ? "sa"
                              : "su"
                          ),
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : `${evt.recurrence.range.endDate}T${
                                  evt.end.dateTime.split("T")[1]
                                }`, // will also accept '20120201'
                        },
                        itemFrom: "PersonalCalendar",
                      }
                    : evt.recurrence &&
                      evt.recurrence.pattern.type == "absoluteMonthly"
                    ? {
                        id: evt.id,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        rrule: {
                          freq: "monthly",
                          interval: evt.recurrence.pattern.interval,
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : evt.recurrence.range.endDate, // will also accept '20120201'
                        },
                        itemFrom: "PersonalCalendar",
                      }
                    : evt.recurrence &&
                      evt.recurrence.pattern.type == "relativeMonthly"
                    ? {
                        id: evt.id,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end:
                          evt.recurrence.range.type == "noEnd" ||
                          evt.recurrence.range.endDate == "0001-01-01"
                            ? `${
                                new Date(
                                  new Date(
                                    evt.recurrence.range.startDate
                                  ).setFullYear(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).getFullYear() + 1
                                  )
                                )
                                  .toISOString()
                                  .split("T")[0]
                              }T${evt.end.dateTime.split("T")[1]}`
                            : `${evt.recurrence.range.endDate}T${
                                evt.end.dateTime.split("T")[1]
                              }`,
                        display: "block",
                        attendees: evt.attendees,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,

                        rrule: {
                          freq: "monthly",
                          interval: evt.recurrence.pattern.interval,
                          // index: evt.recurrence.pattern.index,
                          byweekday: evt.recurrence.pattern.daysOfWeek.map(
                            (day) =>
                              day == "monday"
                                ? "mo"
                                : day == "tuesday"
                                ? "tu"
                                : day == "wednesday"
                                ? "we"
                                : day == "thursday"
                                ? "th"
                                : day == "friday"
                                ? "fr"
                                : day == "saturday"
                                ? "sa"
                                : day == "sunday"
                                ? "su"
                                : ""
                          ),
                          bysetpos:
                            evt.recurrence.pattern.index == "first"
                              ? 1
                              : evt.recurrence.pattern.index == "second"
                              ? 2
                              : evt.recurrence.pattern.index == "third"
                              ? 3
                              : evt.recurrence.pattern.index == "fourth"
                              ? 4
                              : -1,
                          dtstart: `${evt.recurrence.range.startDate}T${
                            evt.start.dateTime.split("T")[1]
                          }`, // will also accept '20120201T103000'
                          until:
                            evt.recurrence.range.type == "noEnd" ||
                            evt.recurrence.range.endDate == "0001-01-01"
                              ? `${
                                  new Date(
                                    new Date(
                                      evt.recurrence.range.startDate
                                    ).setFullYear(
                                      new Date(
                                        evt.recurrence.range.startDate
                                      ).getFullYear() + 1
                                    )
                                  )
                                    .toISOString()
                                    .split("T")[0]
                                }T${evt.end.dateTime.split("T")[1]}`
                              : evt.recurrence.range.endDate, // will also accept '20120201'
                        },
                        //  description: evt.bodyPreview,
                      }
                    : {
                        id: evt.id,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end: evt.end.dateTime,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
                      };
                });
                props.spcontext.web.lists
                  .getByTitle("Events")
                  .items.select(
                    "*,RecurrenceData,ParticipantsPicker/EMail,ParticipantsPicker/Title"
                  )
                  .expand("ParticipantsPicker")
                  .get()
                  .then((res) => {
                    console.log(res);
                    res = res.filter((row) => !row.isDelete);
                    let spCalendarData = [];
                    spCalendarData = res.map((re) => {
                      let recResult;
                      let eventColor = "";
                      let eventType = "";

                      let attendees =
                        re.ParticipantsPicker &&
                        re.ParticipantsPicker.length > 0
                          ? re.ParticipantsPicker.map((users) => ({
                              emailAddress: { name: users.Title },
                            }))
                          : [];
                      let eveDescription = re.Description
                        ? re.Description.replace(/(<([^>]+)>)/gi, "")
                        : "";
                      let eventColorArr = arrColorVar.filter((colLi) => {
                        return re.Title.toLowerCase().includes(
                          colLi.Title.toLowerCase()
                        );
                      });
                      eventColorArr.length > 0
                        ? ((eventColor = eventColorArr[0].HexCode),
                          (eventType = "GroupCalendar"))
                        : ((eventColor = arrColorVar.filter(
                            (colLi) => colLi.DefaultEventColor == true
                          )[0].HexCode),
                          (eventType = "GroupCalendar Other"));
                      var parser = new DOMParser();
                      var eventDescription = parser.parseFromString(
                        re.Description,
                        "text/html"
                      );
                      console.log(eventDescription);
                      let extractedXML = xmlToJSON.parseString(
                        re.RecurrenceData
                      );
                      recResult = extractedXML.recurrence
                        ? extractedXML.recurrence[0].rule[0]
                        : false;
                      return re.fRecurrence &&
                        recResult &&
                        recResult.repeat[0].daily &&
                        recResult.repeat[0].daily[0]._attr.weekday
                        ? {
                            id: re.GUID,
                            title: re.Title,
                            start: re.EventDate,
                            end: re.EndDate,
                            display: "block",
                            attendees: attendees,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            description: eveDescription,
                            allDay: re.fAllDayEvent,
                            itemFrom: eventType,
                            rrule: {
                              freq: "daily",
                              interval: 1,
                              byweekday: re.fRecurrence &&
                                recResult.repeat[0].daily[0]._attr.weekday &&
                                recResult.repeat[0].daily[0]._attr.weekday
                                  ._value && ["mo", "tu", "we", "th", "fr"],
                              dtstart: re.EventDate, // will also accept '20120201T103000'
                              until: re.EndDate, // will also accept '20120201'
                            },
                          }
                        : re.fRecurrence &&
                          recResult &&
                          recResult.repeat[0].daily &&
                          !recResult.repeat[0].daily[0]._attr.weekday
                        ? {
                            id: re.GUID,
                            title: re.Title,
                            start: re.EventDate,
                            end: re.EndDate,
                            display: "block",
                            attendees: attendees,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            description: eveDescription,
                            allDay: re.fAllDayEvent,
                            itemFrom: eventType,
                            rrule: {
                              freq: "daily",
                              interval:
                                recResult.repeat[0].daily[0]._attr.dayFrequency
                                  ._value,
                              byweekday: re.fRecurrence &&
                                recResult.repeat[0].daily[0]._attr.weekday &&
                                recResult.repeat[0].daily[0]._attr.weekday
                                  ._value && ["mo", "tu", "we", "th", "fr"],
                              dtstart: re.EventDate, // will also accept '20120201T103000'
                              until: re.EndDate, // will also accept '20120201'
                            },
                          }
                        : re.fRecurrence &&
                          recResult &&
                          recResult.repeat[0].weekly
                        ? {
                            id: re.GUID,
                            title: re.Title,
                            start: re.EventDate,
                            end: re.EndDate,
                            display: "block",
                            attendees: attendees,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            description: eveDescription,
                            allDay: re.fAllDayEvent,
                            itemFrom: eventType,
                            rrule: {
                              freq: "weekly",
                              interval:
                                recResult.repeat[0].weekly[0]._attr
                                  .weekFrequency._value,
                              byweekday: Object.keys(
                                recResult.repeat[0].weekly[0]._attr
                              ).pop(),
                              dtstart: re.EventDate, // will also accept '20120201T103000'
                              until: re.EndDate, // will also accept '20120201'
                            },
                          }
                        : re.fRecurrence &&
                          recResult &&
                          recResult.repeat[0].monthly
                        ? {
                            id: re.GUID,
                            title: re.Title,
                            start: re.EventDate,
                            end: re.EndDate,
                            display: "block",
                            attendees: attendees,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            description: eveDescription,
                            allDay: re.fAllDayEvent,
                            itemFrom: eventType,
                            rrule: {
                              freq: "monthly",
                              interval:
                                recResult.repeat[0].monthly[0]._attr
                                  .monthFrequency._value,
                              byweekday: Object.keys(
                                recResult.repeat[0].monthly[0]._attr
                              ).pop(),
                              dtstart: re.EventDate, // will also accept '20120201T103000'
                              until: re.EndDate, // will also accept '20120201'
                            },
                          }
                        : re.fRecurrence &&
                          recResult &&
                          recResult.repeat[0].monthlyByDay
                        ? {
                            id: re.GUID,
                            title: re.Title,
                            start: re.EventDate,
                            end: re.EndDate,
                            display: "block",
                            attendees: attendees,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            description: eveDescription,
                            allDay: re.fAllDayEvent,
                            itemFrom: eventType,
                            rrule: {
                              freq: "monthly",
                              interval:
                                recResult.repeat[0].monthlyByDay[0]._attr
                                  .monthFrequency._value,
                              byweekday: [
                                re.fRecurrence &&
                                recResult.repeat[0].monthlyByDay[0].mo
                                  ? "mo"
                                  : re.fRecurrence &&
                                    recResult.repeat[0].monthlyByDay[0].tu
                                  ? "tu"
                                  : recResult.repeat[0].monthlyByDay[0].we
                                  ? "we"
                                  : recResult.repeat[0].monthlyByDay[0].th
                                  ? "th"
                                  : recResult.repeat[0].monthlyByDay[0].fr
                                  ? "fr"
                                  : recResult.repeat[0].monthlyByDay[0].sa
                                  ? "sa"
                                  : "su",
                              ],
                              dtstart: re.EventDate, // will also accept '20120201T103000'
                              until: re.EndDate, // will also accept '20120201'
                              bysetpos:
                                recResult.repeat[0].monthlyByDay[0]._attr
                                  .weekdayOfMonth._value == "first"
                                  ? 1
                                  : recResult.repeat[0].monthlyByDay[0]._attr
                                      .weekdayOfMonth._value == "second"
                                  ? 2
                                  : recResult.repeat[0].monthlyByDay[0]._attr
                                      .weekdayOfMonth._value == "third"
                                  ? 3
                                  : recResult.repeat[0].monthlyByDay[0]._attr
                                      .weekdayOfMonth._value == "fourth"
                                  ? 4
                                  : -1,
                            },
                          }
                        : {
                            id: re.GUID,
                            title: re.Title,
                            start: re.EventDate,
                            end: re.EndDate,
                            display: "block",
                            attendees: attendees,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            description: eveDescription,
                            allDay: re.fAllDayEvent,
                            itemFrom: eventType,
                          };
                    });
                    data = [...data, ...spCalendarData];
                    console.log(data);
                    setEvents(data);
                  });
              });
          });
      });
  };
  // TODO Filtered Data on Click
  allData = arrColor.map((aC) => {
    return aC.Title;
  });

  const getSelectedItem = (selectedItems) => {
    setSelectedEvent(selectedItems);
  };

  let personalEventColor = arrColor.filter((aC) => {
    return aC.IsUser === true;
  })[0];

  let defaultEventColor = arrColor.filter((aC) => {
    return aC.DefaultEventColor === true;
  })[0];

  if (selectedEvent.length > 0) {
    isOnload = false;
    FilteredData = [];
    let currentUserEvents = [];
    let otherEvents = [];
    FilteredData = events.filter((selectedItem) => {
      return (
        selectedEvent.filter((sE) =>
          selectedItem.title.toLowerCase().includes(sE.toLowerCase())
        ).length > 0 && selectedItem.itemFrom == "GroupCalendar"
      );
    });
    if (
      selectedEvent.filter((sE) => sE === personalEventColor.Title).length > 0
    ) {
      currentUserEvents = events.filter(
        (evt) => evt.itemFrom === "PersonalCalendar"
      );
      FilteredData = [...FilteredData, ...currentUserEvents];
    }
    if (
      selectedEvent.filter((sE) => sE === defaultEventColor.Title).length > 0
    ) {
      otherEvents = events.filter(
        (evt) => evt.itemFrom === "GroupCalendar Other"
      );
      FilteredData = [...FilteredData, ...otherEvents];
    }
    if (FilteredData.length === 0) {
      FilteredData = [
        {
          allDay: false,
          attendees: [],
          backgroundColor: "",
          borderColor: "",
          description: "",
          display: "",
          end: "",
          id: "",
          itemFrom: "",
          start: "",
          title: "",
        },
      ];
    }
  } else if (isOnload) {
    FilteredData = data;
  } else {
    FilteredData = [
      {
        allDay: false,
        attendees: [],
        backgroundColor: "",
        borderColor: "",
        description: "",
        display: "",
        end: "",
        id: "",
        itemFrom: "",
        start: "",
        title: "",
      },
    ];
  }
  // TODO Filtered Data on Click
  useEffect(() => {
    // getEventsFromCalendar();
    getEventsFromEvents();
    // Calendar Calls
  }, []);
  return (
    <div className="calendar-section">
      <CalendarColorView
        allData={allData}
        spcontext={props.spcontext}
        arrColor={arrColor}
        onItemClick={getSelectedItem}
      />
      <CalendarDetails
        calendarValue={FilteredData}
        spcontext={props.spcontext}
        graphcontext={props.graphcontext}
      />
    </div>
  );
};
export default CalendarApp;
