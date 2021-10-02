import { ICalendar } from "office-ui-fabric-react";
import TuiCalendar, { ICalendarInfo, IEventDateObject, IEventMoreObject, IEventObject, IEventScheduleObject, IMonthOptions, ISchedule, ITemplateConfig, ITheme, ITimezone, IWeekOptions } from 'tui-calendar';

export interface ICalendarProps {
    description?: string;
    schedules? : ISchedule[];
    defaultView?:string;
    view?:string;
    calendars?:ICalendarInfo[];
    height?:string;
    theme?: ITheme;
    disableDblClick?: boolean;
    disableClick?: boolean;
    isReadOnly?: boolean;
    usageStatistics?: boolean;     
    month?: IMonthOptions;     
    taskView?: boolean | string[];
    scheduleView?: boolean | string[];
    template?: ITemplateConfig;     
    timezones?: ITimezone[];     
    useCreationPopup?: boolean;
    useDetailPopup?: boolean;    
    week?: IWeekOptions;
    'onAfterRenderSchedule'?: (eventObj: {schedule: ISchedule}) => void;
    'onBeforeCreateSchedule'?: (schedule: ISchedule) => void;
    'onBeforeDeleteSchedule'?: (eventObj: IEventScheduleObject) => void;
    'onBeforeUpdateSchedule'?: (eventObj: IEventObject) => void;
    'onClickDayname'?: (eventObj: IEventDateObject) => void;
    'onClickMore'?: (eventObj: IEventMoreObject) => void;
    'onClickSchedule'?: (eventObj: IEventScheduleObject) => void;
    'onClickTimezonesCollapseBtn'?: (timezonesCollapsed: boolean) => void;
  }