import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

import axios from "axios";

/* global console, Excel, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };

    this.columnLabels = ["Name", "Place", "Animal", "Thing"];
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Select",
          primaryText: "Select a four column Table",
        },
        {
          icon: "Click",
          primaryText: "Click Predict button below",
        },
      ],
    });
  }

  predictTable = async () => {
    await Excel.run(async (context) => {
      let range = context.workbook.getSelectedRange();

      range.load("values");
      await context.sync();

      let { data } = await this.predictRequest(range.values);
      let columnNames = [];

      for (let d of data) {
        columnNames.push(this.columnLabels[d]);
      }

      let firstRow = range.getRow(0);
      let headerRow = firstRow.insert(Excel.InsertShiftDirection.down);

      headerRow.values = [columnNames];

      await context.sync();
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  };

  predictRequest = async (values) => {
    let data = JSON.stringify({
      data: values,
    });

    let config = {
      method: "post",
      url: "http://127.0.0.1:5000/predict",
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
      },
      data: data,
    };

    try {
      let res = await axios(config);
      return res.data;
    } catch (err) {
      console.error("ERROR IN REQUEST: ", err);
    }
  };

  /*  {
    "data": [
        0,
        0,
        0,
        0
    ]
} */

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Classifier" />
        <HeroList message="Follow below steps" items={this.state.listItems}>
          <p className="ms-font-l">
            Select 4 column table then click <b>Predict</b>.
          </p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.predictTable}
          >
            Predict
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
