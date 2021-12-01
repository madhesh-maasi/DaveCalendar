import { template } from "lodash";
import * as React from "react";
import { useState, useEffect } from "react";
import CalendarColorItem from "./CalendarColorItem";
import "./CalendarColorView.css";
let isOnLoad = true;
function CalendarColorView(props) {
  const [selectedItems, setSelectedItems] = useState([]);

  const selectedItemHandler = (selectedItem) => {
    selectedItems.length > 0
      ? selectedItems.filter((item) => item == selectedItem).length == 0
        ? setSelectedItems([...selectedItems, selectedItem])
        : setSelectedItems([...selectedItems])
      : setSelectedItems([...selectedItems, selectedItem]);
  };
  const deSelectedItemHandler = (deSelectedItem) => {
    let tempArray = [];
    if (selectedItems.length == 0 && isOnLoad) {
      isOnLoad = false;
      tempArray = props.allData;
    } else {
      tempArray = selectedItems;
    }
    // let tempArray = selectedItems;
    tempArray = tempArray.filter((tA) => tA != deSelectedItem);
    setSelectedItems(tempArray);
  };

  // if (props.allData.length > 0) {
  //   setSelectedItems(props.allData);
  // }

  props.onItemClick(selectedItems);
  return (
    <div className="color-info-section">
      {props.arrColor.map((li) => {
        return (
          <CalendarColorItem
            title={li.Title}
            id={li.ID}
            hex={li.HexCode}
            onSelected={selectedItemHandler}
            onDeSelected={deSelectedItemHandler}
          />
        );
      })}
    </div>
  );
}
export default CalendarColorView;
