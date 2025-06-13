import * as React from "react";
import "./App.css";
import ClearDocument from "./ClearDocument";
import InsertDocument from "./InsertDocument";
import Description from "./Description";

const App = () => {
  return (
    <div className="wrapper">
      <h4>Bug Demo</h4>
      <ClearDocument />
      <InsertDocument />
      <Description />
    </div>
  );
};

export default App;
