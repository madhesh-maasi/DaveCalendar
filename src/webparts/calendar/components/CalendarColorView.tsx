import * as React from "react";
import { useState, useEffect } from "react";
import "./CalendarColorView.css";
function CalendarColorView(props) {
  const [arrColLi, setarrColLi] = useState([]);
  useEffect(() => {
    props.spcontext.web.lists
      .getByTitle("CalColorConfig")
      .items.get()
      .then((data) => {
        setarrColLi(data);
      });
  });
  return (
    <div className="color-info-section">
      {arrColLi.map((li) => {
        return (
          <div className="color-item">
            <div
              className="item-color"
              style={{
                background: `${li.HexCode}`,
                width: "30px",
                height: "30px",
              }}
            ></div>
            <div className="item-name">{li.Title}</div>
          </div>
        );
      })}
    </div>
  );
}
export default CalendarColorView;
