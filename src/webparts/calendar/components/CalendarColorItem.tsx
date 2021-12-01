import * as React from "react";
import { useState, useEffect } from "react";
const CalendarColorItem = (props) => {
  const [isActive, setIsActive] = useState(true);

  return (
    <div
      key={props.id}
      className={`color-item ${isActive && "active"}`}
      style={{
        borderLeft: isActive ? `2px solid ${props.hex}` : `2px solid #fff`,
        cursor: "pointer",
        background: isActive ? `${props.hex}1A` : `#fff`,
      }}
      onClick={() => {
        isActive
          ? props.onDeSelected(props.title)
          : props.onSelected(props.title);
        isActive ? setIsActive(false) : setIsActive(true);
      }}
    >
      <div
        className="item-color"
        style={{
          background: `${props.hex}`,
          width: "30px",
          height: "30px",
        }}
      ></div>
      <div className="item-name">{props.title}</div>
    </div>
  );
};
export default CalendarColorItem;
