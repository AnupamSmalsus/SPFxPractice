export interface ILeaveRequest {
  Id?: number;
  Title: string;
  Employee: string;
  LeaveType: string;
  StartDate: string;
  EndDate: string;
  Reason: string;
  Status: string;
  ManagerComments?: string;
}