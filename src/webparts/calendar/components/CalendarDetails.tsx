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
import rrulePlugin from "@fullcalendar/rrule";
import "./Bootstrap.js";
import "./Bootstrap.css";

let calendar;
let clickedTarget = "";
// let isCalendarFetched = props.canRender;
function CalendarDetails(props) {
  const [ViewItems, setViewItems] = useState({
    Title: "",
    StartDate: "",
    EndDate: "",
    Attendees: "",
    Description: "",
    AllDay: "",
  });

  // let canRenderCal = props.canRender;
  if (props.calendarValue.length > 0) {
    BindCalender(props.calendarValue);
  } else {
    BindCalender([
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
    ]);
  }

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
                <div className="col-6 modalDescription">
                  {ViewItems.Description}
                </div>
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
      plugins: [
        rrulePlugin,
        interactionPlugin,
        dayGridPlugin,
        timeGridPlugin,
        listPlugin,
      ],
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

      showNonCurrentDates: false,
      eventDidMount: (event) => {
        event.el.setAttribute("data-id", event.event.id);
        event.el.setAttribute("data-bs-target", "#viewItemModal");
        event.el.setAttribute("data-bs-toggle", "modal");
        event.el.setAttribute("title", event.event.title);
        event.el.classList.add("view-event");
        // ! Show Event Click
        event.el.addEventListener("click", (e) => {
          clickedTarget = e.target["className"];
          let indexID = event.event.id;
          let viewItem = props.calendarValue.filter(
            (li) => li.id == indexID
          )[0];
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

    if (clickedTarget == "" && calendarEl != null) {
      calendar.render();
      calendar.refetchEvents();
    } else {
      clickedTarget = "";
    }

    let dragClass = document.querySelectorAll(".fc-event-draggable");
    dragClass.forEach((dC) => {
      dC.classList.remove("fc-event-draggable");
      dC.classList.add("view-event");
    });
  }
}

export default CalendarDetails;
