import * as React from 'react';
import { IVenCalendarProps } from './IVenCalendarProps';
import { escape, times } from '@microsoft/sp-lodash-subset';
import TuiCalendar, { ISchedule } from 'tui-calendar';
import 'tui-time-picker/dist/tui-time-picker.css';
import 'tui-date-picker/dist/tui-date-picker.css';
import 'tui-calendar/dist/tui-calendar.css';
import Calendar from './Calendar';
import myTheme from './myTheme';
import './Icon.css';
import './Calendar.css';
import * as moment from 'moment';
import styles from './VenCalendar.module.scss';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
const today = new Date();
const getDate = (type, start, value, operator) => {
  start = new Date(start);
  type = type.charAt(0).toUpperCase() + type.slice(1);

  if (operator === '+') {
    start[`set${type}`](start[`get${type}`]() + value);
  } else {
    start[`set${type}`](start[`get${type}`]() - value);
  }

  return start;
};
export default class VenCalendar extends React.Component<IVenCalendarProps, {}> {
  ref = React.createRef<Calendar>();

  calendarInst = null;

  state = {
    dateRange: '',
    view: 'month',
    viewModeOptions: [
      {
        title: 'Monthly',
        value: 'month'
      },
      {
        title: 'Weekly',
        value: 'week'
      },
      {
        title: 'Daily',
        value: 'day'
      }
    ]
  };

  componentDidMount() {
    this.calendarInst = this.ref.current.getInstance();
    this.setState({view: this.props.view});

    this.setRenderRangeText();
    
  }
  componentDidUpdate(prevProps, prevState) {
    if (this.state.view != prevState.view) {
      this.setRenderRangeText(); 
    }
  }

  onAfterRenderSchedule(res) {
    console.group('onAfterRenderSchedule');
    console.log('Schedule Info : ', res.schedule);
    console.groupEnd();
  }

  onBeforeDeleteSchedule(res) {
    console.group('onBeforeDeleteSchedule');
    console.log('Schedule Info : ', res.schedule);
    console.groupEnd();

    const {id, calendarId} = res.schedule;

    this.calendarInst.deleteSchedule(id, calendarId);
  }

  onChangeSelect(ev) {
    this.setState({view: ev.target.value},()=>{});
    this.setRenderRangeText();
  }

  onClickDayname(res) {
    // view : week, day
    console.group('onClickDayname');
    console.log(res.date);
    console.groupEnd();
  }

  onClickNavi(event) {
    if (event.target.tagName === 'BUTTON'||event.target.tagName === 'I') {
      const {target} = event;
      let action = target.dataset ? target.dataset.action : target.getAttribute('data-action');
      action = action.replace('move-', '');

      this.calendarInst[action]();
      this.setRenderRangeText();
    }
  }

  onClickSchedule(res) {
    console.group('onClickSchedule');
    console.log('MouseEvent : ', res.event);
    console.log('Calendar Info : ', res.calendar);
    console.log('Schedule Info : ', res.schedule);
    console.groupEnd();
  }

  onClickTimezonesCollapseBtn(timezonesCollapsed) {
    // view : week, day
    console.group('onClickTimezonesCollapseBtn');
    console.log('Is Collapsed Timezone? ', timezonesCollapsed);
    console.groupEnd();

    const theme = {};
    if (timezonesCollapsed) {
      theme['week.daygridLeft.width'] = '200px';
      theme['week.timegridLeft.width'] = '200px';
    } else {
      theme['week.daygridLeft.width'] = '100px';
      theme['week.timegridLeft.width'] = '100px';
    }

    this.calendarInst.setTheme(theme);
  }

  setRenderRangeText() {
    const view = this.calendarInst.getViewName();
    const calDate = this.calendarInst.getDate();
    const rangeStart = this.calendarInst.getDateRangeStart();
    const rangeEnd = this.calendarInst.getDateRangeEnd();
    
    let year = calDate.getFullYear();
    let month = calDate.getMonth() + 1;
    let date = calDate.getDate();
    let dateRangeText = '';
    let endMonth, endDate, start, end;

    switch (view) {
      case 'month':
        dateRangeText = `${month}-${year}`;
        break;
      case 'week':
        year = rangeStart.getFullYear();
        month = rangeStart.getMonth() + 1;
        date = rangeStart.getDate();
        endMonth = rangeEnd.getMonth() + 1;
        endDate = rangeEnd.getDate();

       /// start = `${year}-${month < 10 ? '0' : ''}${month}-${date < 10 ? '0' : ''}${date}`;
        start = `${date < 10 ? '0' : ''}${date}-${month < 10 ? '0' : ''}${month}-${year}`;
        end =  `${
          endDate < 10 ? '0' : ''
        }${endDate}-${endMonth < 10 ? '0' : ''}${endMonth}-${year}`;
       /* end = `${year}-${endMonth < 10 ? '0' : ''}${endMonth}-${
          endDate < 10 ? '0' : ''
        }${endDate}`;*/
        dateRangeText = `${start} ~ ${end}`;
        break;
      default:
        dateRangeText = `${date}-${month}-${year}`;
    }

    this.setState({dateRange: dateRangeText});
  }

  onBeforeUpdateSchedule(event) {
    const {schedule} = event;
    const {changes} = event;

    this.calendarInst.updateSchedule(schedule.id, schedule.calendarId, changes);
  }

  onBeforeCreateSchedule(scheduleData) {
    const {calendar} = scheduleData;
    const schedule :ISchedule = {
      id: String(Math.random()),
      title: scheduleData.title,
      isAllDay: scheduleData.isAllDay,
      start: scheduleData.start,
      end: scheduleData.end,
      category: scheduleData.isAllDay ? 'allday' : 'time',
      dueDateClass: '',
      location: scheduleData.location,
      raw: {
        class: scheduleData.raw['class']
      },
      state: scheduleData.state
    };

    if (calendar) {
      schedule.calendarId = calendar.id;
      schedule.color = calendar.color;
      schedule.bgColor = calendar.bgColor;
      schedule.borderColor = calendar.borderColor;
    }

    this.calendarInst.createSchedules([schedule]);
  }

  render() {
    const {dateRange, view, viewModeOptions} = this.state;
    const selectedView = view || this.props.view;

    return (
      <div>
        <h1>{this.props.description}</h1>
        <div>
          <select id="calendarTypeName" className="btn btn-default btn-sm dropdown-toggle" onChange={this.onChangeSelect.bind(this)} value={view}>
            {viewModeOptions.map((option, index) => (
              <option value={option.value} key={index}>
                {option.title}
              </option>
            ))}
          </select>
          <span>
            <button
              type="button"
              className="btn btn-default btn-sm move-today"
              data-action="move-today"
              onClick={this.onClickNavi.bind(this)}
            >
              Today
            </button>
            <button
              type="button"
              className="btn btn-default btn-sm move-day"
              data-action="move-prev"
              onClick={this.onClickNavi.bind(this)}
            >
              <i className="calendar-icon ic-arrow-line-left" data-action="move-prev" ></i>
            </button>
            <button
              type="button"
              className="btn btn-default btn-sm move-day"
              data-action="move-next"
              onClick={this.onClickNavi.bind(this)}
            >
              <i className="calendar-icon ic-arrow-line-right" data-action="move-next"></i>
            </button>
          </span>
          <span className="render-range">{dateRange}</span>
        </div>
        <Calendar
          usageStatistics={false}
          calendars={this.props.calendars}
          defaultView={selectedView}
          disableDblClick={true}
          height="900px"
          isReadOnly={false}
          month={{
            startDayOfWeek: 0
          }}
          schedules={this.props.schedules}
          scheduleView = {true}
          taskView
          template={{
            milestone(schedule) {
              return `<span style="color:#fff;background-color: ${schedule.bgColor};">${
                schedule.title
              }</span>`;
            },
            milestoneTitle() {
              return 'Milestone';
            },
            allday(schedule) {
              var calCategories = schedule.raw["calCategories"];
              var calCategory = schedule.raw["CalCategory"]== null?"":schedule.raw["CalCategory"];
              var scheduleAnchorAttri = "";
              var color = (calCategories[calCategory as string]?calCategories[calCategory as string]:"silver");
              if (Environment.type == EnvironmentType.ClassicSharePoint){
                scheduleAnchorAttri = 'onClick="'+schedule.raw["scheduleAnchorAttri"]+'"';
              }else{
                scheduleAnchorAttri = 'href="'+schedule.raw["scheduleAnchorAttri"]+'"';
              }
              return '<a style="display:block;color:#fff;background-color:'+color+'; border-color:'+color+';text-decoration:none" '+scheduleAnchorAttri+' title="'+'All Day'+'\n'+ schedule.title+'\n'+calCategory+'">'+schedule.title+'<i class="fa fa-refresh"></i></a>';
            
              //return `${schedule.title}<i class="fa fa-refresh"></i>`;
            },
            alldayTitle() {
              return 'All Day';
            },
            popupDetailDate: function(isAllDay, start, end) {
              var isSameDate = moment(start as Date).isSame(end as Date);
              var endFormat = (isSameDate ? '' : 'YYYY.MM.DD ') + 'hh:mm a';
        
              if (isAllDay) {
                return moment(start as Date).format('YYYY.MM.DD') + (isSameDate ? '' : ' - ' + moment(end as Date).format('YYYY.MM.DD'));
              }
        
              return (moment(start as Date).format('YYYY.MM.DD hh:mm a') + ' - ' + moment(end as Date).format(endFormat));
            },
            popupDetailLocation: function(schedule) {
              return 'Location : ' + schedule.location;
            },
            popupDetailRepeat: function(schedule) {
                return schedule.recurrenceRule;
            },
            popupDetailBody: function(schedule) {
                return schedule.body;
            },
            time: function(schedule) {
              var calCategories = schedule.raw["calCategories"];
              var calCategory =schedule.raw["CalCategory"]== null?"":schedule.raw["CalCategory"];
              var color = (calCategories[calCategory as string]?calCategories[calCategory as string]:"silver");
              var scheduleAnchorAttri = "";
              if (Environment.type == EnvironmentType.ClassicSharePoint){
                scheduleAnchorAttri = 'onClick="'+schedule.raw["scheduleAnchorAttri"]+'"';
              }else{
                scheduleAnchorAttri = 'href="'+schedule.raw["scheduleAnchorAttri"]+'"';
              }
             return '<a style="display:block;color:#fff;background-color:'+color+'; border-color:'+color+'; text-decoration:none" '+scheduleAnchorAttri+' title="'+moment((schedule.start as Date).getTime()).format('HH:mm')+' - '+moment((schedule.end as Date).getTime()).format('HH:mm')+'\n'+ schedule.title+'\n'+calCategory+'">'+moment((schedule.start as Date).getTime()).format('HH:mm') + ' <i class="fa fa-refresh"></i>' + schedule.title+'</a>';
            },
            popupEdit: function() {
              return "";
            },
            popupDelete: function() {
                return "";
            }
          }}
          theme={myTheme}
          useDetailPopup = {false}
          useCreationPopup= {false}
          view={selectedView}
          week={{
            showTimezoneCollapseButton: true,
            timezonesCollapsed: false
          }}
          ref={this.ref}
          onAfterRenderSchedule={this.onAfterRenderSchedule.bind(this)}
          onBeforeDeleteSchedule={this.onBeforeDeleteSchedule.bind(this)}
          onClickDayname={this.onClickDayname.bind(this)}
          onClickSchedule={this.onClickSchedule.bind(this)}
          onClickTimezonesCollapseBtn={this.onClickTimezonesCollapseBtn.bind(this)}
          onBeforeUpdateSchedule={this.onBeforeUpdateSchedule.bind(this)}
          onBeforeCreateSchedule={this.onBeforeCreateSchedule.bind(this)}
        />
      </div>
    );
    
  }
}
