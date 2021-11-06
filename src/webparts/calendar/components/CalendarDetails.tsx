import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./Calendar.module.scss";
import { MSGraphClient } from "@microsoft/sp-http";
import { graph } from "@pnp/graph";
import { Calendar } from "@fullcalendar/core";
import interactionPlugin from "@fullcalendar/interaction";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import listPlugin from "@fullcalendar/list";
import "./Bootstrap.js";
import "./Bootstrap.css";
import ReactTooltip from "react-tooltip";
import { lowerFirst } from "lodash";
// import Moment from 'react-moment';
// import 'moment-timezone';
// Latest to Kamesh
let calendar;
let data = [];
let arrColor = [];
let userInGroup = false;
let calendarLoadFirst = true;
let timeZone = "Pacific Standard Time";
let headers = { Prefer: 'outlook.timezone="' + timeZone + '"' };
function CalendarDetails(props) {
  const [events, setevents] = useState([]);
  const [load, setload] = useState("");
  const [ViewItems, setViewItems] = useState({
    Title: "",
    StartDate: "",
    EndDate: "",
    Attendees: "",
    Description: "",
    AllDay: "",
  });

  if (events.length > 0) {
    BindCalender(events);
  }

  useEffect(() => {
    props.spcontext.web.lists
      .getByTitle("CalMonthConfig")
      .items.get()
      .then(async (li) => {
        await props.spcontext.web.lists
          .getByTitle("CalColorConfig")
          .items.get()
          .then((ccLi) => {
            arrColor = ccLi;
          });
        let date = new Date();
        let firstDay: any = new Date(date.getFullYear(), date.getMonth(), 1);
        let lastDay: any = new Date(date.getFullYear(), date.getMonth() + 1, 0);

        firstDay.setMonth(firstDay.getMonth() - li[0].Month);
        lastDay.setMonth(lastDay.getMonth() + li[0].Month);

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
        if (li[0].GroupID != null) {
          await graph.groups
            .getById(li[0].GroupID)
            .members.top(999)()
            .then((groupRes: any) => {
              userInGroup = groupRes.filter((gR) => gR.id == myId).length > 0;
            });
        }

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
              let myEventColor = arrColor.filter((aC) => aC.IsUser == true)[0]
                .HexCode;
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
                  })
                : "";
              return evt.recurrence && evt.recurrence.pattern.type == "weekly"
                ? {
                    id: evt.id,
                    daysOfWeek: dow,
                    startRecur: evt.recurrence.range.startDate,
                    endRecur: evt.recurrence.range.endDate,
                    title: evt.subject,
                    start: evt.start.dateTime,
                    end: evt.end.dateTime,
                    display: "block",
                    attendees: evt.attendees,
                    backgroundColor: myEventColor,
                    borderColor: myEventColor,
                    description: evt.bodyPreview,
                    allDay: evt.isAllDay,
                  }
                : evt.recurrence && evt.recurrence.pattern.type == "daily"
                ? {
                    id: evt.id,
                    // daysOfWeek: [1, 2, 3, 4, 5, 6, 7],
                    startRecur: evt.recurrence.range.startDate,
                    endRecur: evt.recurrence.range.endDate,
                    title: evt.subject,
                    start: evt.start.dateTime,
                    end: evt.end.dateTime,
                    display: "block",
                    attendees: evt.attendees,
                    backgroundColor: myEventColor,
                    borderColor: myEventColor,
                    description: evt.bodyPreview,
                    allDay: evt.isAllDay,
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
                  };
            });
          })
          .then(async () => {
            userInGroup && li[0].GroupID != null
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
                    data1 = result.map((evt) => {
                      let eventColor = "";
                      let eventColorArr = arrColor.filter((colLi) => {
                        return evt.subject
                          .toLowerCase()
                          .includes(colLi.Title.toLowerCase());
                      });
                      eventColorArr.length > 0
                        ? (eventColor = eventColorArr[0].HexCode)
                        : (eventColor = arrColor.filter(
                            (colLi) => colLi.DefaultEventColor == true
                          )[0].HexCode);

                      console.log(eventColor);

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
                          })
                        : "";
                      return evt.recurrence &&
                        evt.recurrence.pattern.type == "weekly"
                        ? {
                            id: evt.id,
                            title: evt.subject,
                            daysOfWeek: dow,
                            startRecur: evt.recurrence.range.startDate,
                            endRecur: evt.recurrence.range.endDate,
                            start: evt.start.dateTime,
                            end: evt.end.dateTime,
                            display: "block",
                            attendees: evt.attendees,
                            description: evt.bodyPreview,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            allDay: evt.isAllDay,
                            //  description: evt.bodyPreview,
                          }
                        : evt.recurrence &&
                          evt.recurrence.pattern.type == "daily"
                        ? {
                            id: evt.id,
                            title: evt.subject,
                            daysOfWeek: [1, 2, 3, 4, 5, 6, 7],
                            startRecur: evt.recurrence.range.startDate,
                            endRecur: evt.recurrence.range.endDate,
                            start: evt.start.dateTime,
                            end: evt.end.dateTime,
                            display: "block",
                            attendees: evt.attendees,
                            description: evt.bodyPreview,
                            backgroundColor: eventColor,
                            borderColor: eventColor,
                            allDay: evt.isAllDay,
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
                            //  description: evt.bodyPreview,
                          };
                    });
                    data = [...data, ...data1];
                    setevents(data);
                  })
              : setevents(data);
            // console.log(events);
          });
      });
  }, []);
  return (
    <div className="w-100">
      <div className="calendar-section" id="myCalendar"></div>

      <button
        type="button"
        className="btn btn-primary btn-open-view d-none"
        data-bs-toggle="modal"
        data-bs-target="#viewItemModal"
      >
        Open Modal
      </button>

      <div
        className="modal fade"
        id="viewItemModal"
        data-bs-backdrop="static"
        data-bs-keyboard="false"
        aria-labelledby="viewItemModalLabel"
        aria-hidden="true"
      >
        <div className="modal-dialog">
          <div className="modal-content">
            <div className="modal-header">
              <h5 className="modal-title m-auto" id="viewItemModalLabel">
                View Event
              </h5>
            </div>
            <div className="modal-body">
              <div className="row my-3">
                <div className="col-5 modal-label">Title</div>
                <div className="col-1">:</div>
                <div className="col-6">{ViewItems.Title}</div>
              </div>
              <div className="row my-3">
                <div className="col-5 modal-label">Start Date and Time</div>
                <div className="col-1">:</div>
                <div className="col-6">{ViewItems.StartDate}</div>
              </div>
              <div className="row my-3">
                <div className="col-5 modal-label">End Date and Time</div>
                <div className="col-1">:</div>
                <div className="col-6">{ViewItems.EndDate}</div>
              </div>
              <div className="row my-3">
                <div className="col-5 modal-label">Attendees</div>
                <div className="col-1">:</div>
                <div className="col-6">{ViewItems.Attendees}</div>
              </div>
              <div className="row my-3">
                <div className="col-5 modal-label">All Day</div>
                <div className="col-1">:</div>
                <div className="col-6">{ViewItems.AllDay}</div>
              </div>
              <div className="row my-3">
                <div className="col-5 modal-label">Description</div>
                <div className="col-1">:</div>
                <div className="col-6">{ViewItems.Description}</div>
              </div>
            </div>
            <div className="modal-footer">
              <button
                type="button"
                className="btn btn-secondary"
                data-bs-dismiss="modal"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
  function BindCalender(Calendardetails) {
    // calendar.refetchEvents();
    // !Calendar Bind
    const dateFormate = new Date("1976-04-19T12:59-0500");
    var calendarEl = document.getElementById("myCalendar");
    calendar = new Calendar(calendarEl, {
      plugins: [interactionPlugin, dayGridPlugin, timeGridPlugin, listPlugin],
      headerToolbar: {
        left: "prev,next today",
        center: "title",
        right: "dayGridMonth",
      },
      initialDate: new Date(),
      navLinks: true, // can click day/week names to navigate views
      editable: true,
      dayMaxEvents: true, // allow "more" link when too many events
      events: Calendardetails,
      eventDidMount: (event) => {
        event.el.setAttribute("data-id", event.event.id);
        event.el.setAttribute("data-bs-target", "#viewItemModal");
        event.el.setAttribute("data-bs-toggle", "modal");
        event.el.setAttribute("title", event.event.title);
        event.el.classList.add("view-event");
        // ! Show Event Click
        event.el.addEventListener("click", (e) => {
          let indexID = event.event.id;
          let viewItem = data.filter((li) => li.id == indexID)[0];
          // console.log(viewItem);
          let attendees = "";
          if (viewItem.attendees.length > 0) {
            viewItem.attendees.forEach((att) => {
              attendees += `${att.emailAddress.name}; `;
            });
          }
          setViewItems({
            Title: viewItem.title,
            StartDate: new Date(viewItem.start).toLocaleString().toString(),
            EndDate: new Date(viewItem.end).toLocaleString().toString(),
            Attendees: attendees,
            Description: viewItem.description,
            AllDay: viewItem.allDay ? "Yes" : "No",
          });
        });
      },
    });
    // ! Locked Rerender of Calendar
    if (calendarLoadFirst) {
      calendar.refetchEvents();
      calendar.render();
      calendarLoadFirst = false;
    }

    let dragClass = document.querySelectorAll(".fc-event-draggable");
    dragClass.forEach((dC) => {
      dC.classList.remove("fc-event-draggable");
      dC.classList.add("view-event");
    });
  }
}

export default CalendarDetails;
