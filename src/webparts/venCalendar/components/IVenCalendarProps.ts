import { ICalendar } from "office-ui-fabric-react";
import TuiCalendar, { DateType, ICalendarInfo, IGridDateModel, IMonthDayNameInfo, ISchedule, ITheme, ITimeGridHourLabel, ITimezoneHourMarker, IWeekDayNameInfo } from 'tui-calendar';

export interface IVenCalendarProps {
  description?: string;
  schedules? : ISchedule[];
  view?:string;
  calendars?:ICalendarInfo[];
  height?:string;
  theme?: ITheme; 
  calCategories? : {};
}

