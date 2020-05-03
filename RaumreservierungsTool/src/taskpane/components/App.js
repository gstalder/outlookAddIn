import * as React from "react";
import { Button, ButtonType, PrimaryButton, DefaultButton } from "office-ui-fabric-react";
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings, mergeStyleSets } from 'office-ui-fabric-react';
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
/* global Button, Header, HeroList, HeroListItem, Progress */

// Used to add spacing between example checkboxes
const stackTokens = { childrenGap: 10 };
const horizontalStackTockens = { childrenGap: 50 };
const seatingOptions = [
  { key: 'theatre', text: 'Theater' },
  { key: 'uform', text: 'U-Form' },
  { key: 'education', text: 'Schulung' },
];

const DayPickerStrings = {
  months: [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker',
};

const timeOptions = [
  { key: "00:00", text: "00:00" },
  { key: "00:30", text: "00:30" },
  { key: "01:00", text: "01:00" },
  { key: "01:30", text: "01:30" },
  { key: "02:00", text: "02:00" },
  { key: "02:30", text: "02:30" },
  { key: "03:00", text: "03:00" },
  { key: "03:30", text: "03:30" },
  { key: "04:00", text: "04:00" },
  { key: "04:30", text: "04:30" },
  { key: "05:00", text: "05:00" },
  { key: "05:30", text: "05:30" },
  { key: "06:00", text: "06:00" },
  { key: "06:30", text: "06:30" },
  { key: "07:00", text: "07:00" },
  { key: "07:30", text: "07:30" },
  { key: "08:00", text: "08:00" },
  { key: "08:30", text: "08:30" },
  { key: "09:00", text: "09:00" },
  { key: "09:30", text: "09:30" },
  { key: "10:00", text: "10:00" },
  { key: "10:30", text: "10:30" },
  { key: "11:00", text: "11:00" },
  { key: "11:30", text: "11:30" },
  { key: "12:00", text: "12:00" },
  { key: "12:30", text: "12:30" },
  { key: "13:00", text: "13:00" },
  { key: "13:30", text: "13:30" },
  { key: "14:00", text: "14:00" },
  { key: "14:30", text: "14:30" },
  { key: "15:00", text: "15:00" },
  { key: "15:30", text: "15:30" },
  { key: "16:00", text: "16:00" },
  { key: "16:30", text: "16:30" },
  { key: "17:00", text: "17:00" },
  { key: "17:30", text: "17:30" },
  { key: "18:00", text: "18:00" },
  { key: "18:30", text: "18:30" },
  { key: "19:00", text: "19:00" },
  { key: "19:30", text: "19:30" },
  { key: "20:00", text: "20:00" },
  { key: "20:30", text: "20:30" },
  { key: "21:00", text: "21:00" },
  { key: "21:30", text: "21:30" },
  { key: "22:00", text: "22:00" },
  { key: "22:30", text: "22:30" },
  { key: "23:00", text: "23:00" },
  { key: "23:30", text: "23:30" },
];

const controlClass = mergeStyleSets({
  control: {
    margin: '0 0 15px 0',
    maxWidth: '300px',
  },
});

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      roomId: "",
      startDate: new Date(),
      endDate: new Date(),
      numberOfPeople: "",
      seating: "",
      food: false,
      audio: false,
      video: false
    };
  }

  onChangeNumberOfPeople(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, typeof data);
      this.setState({ numberOfPeople: data });
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  onChangeSeating(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data.key, typeof data.key);
      this.setState({ seating: data.key });
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  onChangeFood(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, typeof data);
      this.setState({ food: data });
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  onChangeAudio(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, typeof data);
      this.setState({ audio: data });
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  onChangeVideo(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, typeof data);
      this.setState({ video: data });
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  onChangeDatePicker(data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, typeof data);
      var date = this.composetDate(data);
      this.setState({ startDate: date });
      this.setState({ endDate: date });
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  composetDate(date) {
    var returnDate = this.state.startDate;
    returnDate.setFullYear(date.getFullYear());
    returnDate.setMonth(date.getMonth());
    returnDate.setDate(date.getDate());
    console.log(returnDate);
    return returnDate;
  };

  onChangeStartTime(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, typeof data);
      if (this.validateInput(data.key)) {
        this.composeStartTime(data.key.replace(/\D.*/g, ""), data.key.replace(/.*\D/g, ""));
      };
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  composeStartTime(hours, minutes) {
    this.state.startDate.setHours(Number(hours));
    this.state.startDate.setMinutes(Number(minutes));
  };

  onChangeEndTime(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, typeof data);
      if (this.validateInput(data.key)) {
        this.composeEndTime(data.key.replace(/\D.*/g, ""), data.key.replace(/.*\D/g, ""));
      };
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  composeEndTime(hours, minutes) {
    this.state.endDate.setHours(Number(hours));
    this.state.endDate.setMinutes(Number(minutes));
  };

  validateInput(str) {
    let validRegEx = /\d\d:\d\d/;
    return str.match(validRegEx) === null ? false : true;
  };

  validateNumberOfPeople(value) {
    if (value === null ||
      value.trim().length === 0
      || value === "__") {
      return 'Bitte angeben angeben';
    }
    return "";
  };

  onChangeRoomId(evt, data) {
    var printError = function (error, explicit) {
      console.log(`[${explicit ? 'EXPLICIT' : 'INEXPLICIT'}] ${error.name}: ${error.message}`);
    }
    try {
      console.log(data, " ", typeof data, " ", data.length);
      const extract = (str, pattern) => (str.match(pattern) || []).pop() || '';
      const extractAlphanum = (str) => extract(str, "[0-9a-zA-Z]+");
      console.log("alphanum: ", extractAlphanum(data));
      this.setState({ roomId: extractAlphanum(data) });
    } catch (e) {
      if (e instanceof TypeError) {
        printError(e, true);
      } else {
        printError(e, false);
      }
    }
  };

  copyToBody = async () => {
    if (this.validateNumberOfPeople(this.state.numberOfPeople) != "") return;
    var content;
    this.setStartTime();
    this.setEndTime();
    content = this.prepareContent();
    Office.context.mailbox.item.body.setSelectedDataAsync(content,
      { coercionType: Office.CoercionType.Html }, function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Could not insert options to message body.");
        }
      });
    console.log("Options inserted.");
  };

  setStartTime() {
    var startDate = this.state.startDate;
    Office.context.mailbox.item.start.setAsync(
      startDate,
      { asyncContext: { var1: 1, var2: 2 } },
      function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          console.log("Failed to set start date.");
        }
        else {
          console.log("Succesfully set start date.")
        }
      });
  }

  setEndTime() {
    var endDate = this.state.endDate;
    Office.context.mailbox.item.end.setAsync(
      endDate,
      { asyncContext: { var1: 1, var2: 2 } },
      function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          console.log("Failed to set end date.");
        }
        else {
          console.log("Succesfully set end date.")
        }
      });
  }


  sendToServer = async () => {
    this.state.startDate.setSeconds(0);
    this.state.endDate.setSeconds(0);
    // this.state.startDate.setMilliseconds(0);
    // this.state.endDate.setMilliseconds(0);
    var jsonData = JSON.stringify(this.state);
    $.ajax({
      type: "POST",
      url: "http://localhost:8080/json",
      data: jsonData,
      success: function (text) {
        console.log("Success connecting to Server: " + text)
      },
      failure: function (errMsg) {
        console.log(errMsg);
        alert(errMsg);
      }
    });
    Office.context.ui.closeContainer();
  }

  render() {
    const { title, isOfficeInitialized } = this.props;
    const { roomId, startDate } = this.state;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Lädt, bitte warten." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Willkommen" />
        <HeroList items={[]}>
          <p className="ms-font-l">
            Raumeinstellungen angeben, anschliessend <b>an Umsystem senden</b>.
          </p>
          <Stack tokens={stackTokens}>
            <TextField
              label="Raum-ID"
              value={this.state.roomId}
              onChange={this.onChangeRoomId.bind(this)}
              required
            >
            </TextField >
            <DatePicker
              label="Datum und Zeit"
              value={startDate}
              className={controlClass.control}
              strings={DayPickerStrings}
              showWeekNumbers={true}
              firstWeekOfYear={1}
              showMonthPickerAsOverlay={true}
              placeholder="Select a date..."
              ariaLabel="Select a date"
              onSelectDate={this.onChangeDatePicker.bind(this)}
            />
            <Stack horizontal tokens={horizontalStackTockens}>
              <Dropdown
                placeHolder="Start"
                options={timeOptions}
                onChange={this.onChangeStartTime.bind(this)}
              >
              </Dropdown>
              <Dropdown
                value={roomId}
                placeHolder="Ende"
                options={timeOptions}
                onChange={this.onChangeEndTime.bind(this)}
              >
              </Dropdown>
            </Stack>
            <MaskedTextField
              label="Anzahl Personen"
              onChange={this.onChangeNumberOfPeople.bind(this)}
              required
              mask="99"
              onGetErrorMessage={this.validateNumberOfPeople.bind(this)}
            >
            </MaskedTextField >
            <Dropdown
              label="Bestuhlung"
              placeholder="Auswählen"
              options={seatingOptions}
              onChange={this.onChangeSeating.bind(this)}
            >
            </Dropdown>
            <Checkbox
              label="Verpflegung"
              onChange={this.onChangeFood.bind(this)}
            >
            </Checkbox>
            <Checkbox
              label="Audio"
              onChange={this.onChangeAudio.bind(this)}
            >
            </Checkbox>
            <Checkbox
              label="Video"
              onChange={this.onChangeVideo.bind(this)}
            >
            </Checkbox>
            <DefaultButton
              className="ms-welcome__action"
              buttonType={ButtonType.hero}
              onClick={this.copyToBody}
            >
              In Termin kopieren
            </DefaultButton>
            <PrimaryButton
              className="ms-welcome__action"
              buttonType={ButtonType.hero}
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.sendToServer}
            >
              An Umsystem senden
            </PrimaryButton>
          </Stack>
        </HeroList>
      </div>
    );
  }

  prepareContent() {
    var content = "<i>Raumeinstellungen</i><p><table cellpadding=\"2\" cellspacing=\"2\" border=\"0\">";
    for (let key in this.state) {
      if (this.state.hasOwnProperty(key)) {
        if (key !== "endDate") content += this.newHtmlTableEntry(key, this.state[key]);
      }
    }
    content += "</table>";
    return content;
  }

  newHtmlTableEntry(key, value) {
    var newHtmlTableEntry = "<tr><td>"
    newHtmlTableEntry += this.getGermanText(key);
    newHtmlTableEntry += ":<td><i>";
    newHtmlTableEntry += this.getStringValue(key, value);
    return newHtmlTableEntry + "</i></td></tr>";

  }

  getStringValue(key, value) {
    if (key === "seating") {
      var returnVal = seatingOptions.find(obj => {
        return obj.key === value
      })
      return typeof returnVal == "undefined" ? "Bestuhlung nicht definiert" : returnVal.text;
    }
    else if (typeof value === "boolean") {
      return value ? "Ja" : "Nein";
    }
    else if (typeof value === "string") {
      return value.replace("_", "");
    }
    else if (typeof value === "object" && typeof value.getMonth === "function") {
      console.log(typeof value, " ", value);
      var jsonDate = value.toJSON();
      console.log(jsonDate);
      var day = value.getDate();
      var month = value.getMonth() + 1;
      var year = value.getFullYear();
      var startMinutes = this.state.startDate.getMinutes() === 0 ? "00" : this.state.startDate.getMinutes().toString();
      var endMinutes = this.state.endDate.getMinutes() === 0 ? "00" : this.state.endDate.getMinutes().toString();
      var startTime = (this.state.startDate.getHours()) + ":" + startMinutes;
      var endTime = (this.state.endDate.getHours()) + ":" + endMinutes;
      return startTime + " - " + endTime + ", " + day + "." + month + "." + year;
    }
    else {
      if (typeof value === "undefined") {
        return "Nicht definiert";
      }
      return value;
    }
  }

  getGermanText(textKey) {
    if (textKey === "numberOfPeople") {
      return "Anzahl Personen";
    }
    else if (textKey === "food") {
      return "Verpflegung";
    }
    else if (textKey === "audio") {
      return "Audio";
    }
    else if (textKey === "video") {
      return "Video";
    }
    else if (textKey === "seating") {
      return "Bestuhlung";
    }
    else if (textKey === "startDate") {
      return "Datum";
    }
    else if (textKey === "roomId") {
      return "Raum-ID";
    }
    else {
      return "Nicht definiert";
    }
  }
}

