import * as React from "react";
import { useState, useEffect } from "react";
import CalendarColorView from "./CalendarColorView";
import CalendarDetails from "./CalendarDetails";
import { graph } from "@pnp/graph";

let data = [];
let allData = [];
let FilteredData = [];
let arrColorVar = [];
let userInGroup = false;
let timeZone = "Pacific Standard Time"; // For Live
// let timeZone = "Indian Standard Time"; //for local time zone
let headers = { Prefer: 'outlook.timezone="' + timeZone + '"' };
let isOnload = true;
const CalendarApp = (props) => {
  const [arrColor, setArrColor] = useState([]);
  const [events, setEvents] = useState([]);
  const [selectedEvent, setSelectedEvent] = useState([]);
  //   useState
  // colorList calls
  useEffect(() => {
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
                        daysOfWeek: dow,
                        startRecur: evt.recurrence.range.startDate,
                        endRecur: recED,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end: recEndDateTime,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
                      }
                    : evt.recurrence && evt.recurrence.pattern.type == "daily"
                    ? {
                        id: evt.id,
                        // daysOfWeek: [1, 2, 3, 4, 5, 6, 7],
                        startRecur: evt.recurrence.range.startDate,
                        endRecur: recED,
                        title: evt.subject,
                        start: evt.start.dateTime,
                        end: recEndDateTime,
                        display: "block",
                        attendees: evt.attendees,
                        backgroundColor: myEventColor,
                        borderColor: myEventColor,
                        description: evt.bodyPreview,
                        allDay: evt.isAllDay,
                        itemFrom: "PersonalCalendar",
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
                                        daysOfWeek: dow,
                                        startRecur:
                                          evt.recurrence.range.startDate,
                                        endRecur: recED,
                                        start: evt.start.dateTime,
                                        end: recEndDateTime,
                                        display: "block",
                                        attendees: evt.attendees,
                                        description: evt.bodyPreview,
                                        backgroundColor: eventColor,
                                        borderColor: eventColor,
                                        allDay: evt.isAllDay,
                                        itemFrom: eventType,
                                        //  description: evt.bodyPreview,
                                      }
                                    : evt.recurrence &&
                                      evt.recurrence.pattern.type == "daily"
                                    ? {
                                        id: evt.id,
                                        title: evt.subject,
                                        daysOfWeek: [1, 2, 3, 4, 5, 6, 7],
                                        startRecur:
                                          evt.recurrence.range.startDate,
                                        endRecur: recED,
                                        start: evt.start.dateTime,
                                        end: recEndDateTime,
                                        display: "block",
                                        attendees: evt.attendees,
                                        description: evt.bodyPreview,
                                        backgroundColor: eventColor,
                                        borderColor: eventColor,
                                        allDay: evt.isAllDay,
                                        itemFrom: eventType,
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
                                data = [...data, ...data1];
                                setEvents(data);
                              })
                          : "";
                      })
                  : setEvents(data);
              });
          });
      });
    // Calendar Calls
  }, []);

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
