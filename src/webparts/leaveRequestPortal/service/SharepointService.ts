import { SPFI, spfi, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { ILeaveRequest } from "../components/ILeaveRequest";

export class SharePointService {

  private sp: SPFI;

  constructor(context: WebPartContext) {
 
     this.sp = spfi()
       .using(
         SPFx(context)
       );
   }

  /**
   * Get all leave requests
   */
  public async getLeaveRequests(): Promise<ILeaveRequest[]> {

    try {

      const items: any[] = await this.sp.web.lists
        .getByTitle("LeaveRequest")
        .items
        .select(
          "Id",
          "Title",
          "Employee",
          "LeaveType",
          "StartDate",
          "EndDate",
          "Reason",
          "Status",
          "ManagerComments"
        )
        .orderBy("Created", false)
        .top(5000)();

      return items.map((item) => ({
        Id: item.Id,
        Title: item.Title || "",
        Employee: item.Employee || "",
        LeaveType: item.LeaveType || "",
        StartDate: item.StartDate,
        EndDate: item.EndDate,
        Reason: item.Reason || "",
        Status: item.Status || "Pending",
        ManagerComments: item.ManagerComments || ""
      }));

    } catch (error) {

      console.error(
        "Error while fetching leave requests:",
        error
      );

      throw error;
    }
  }


  /**
   * Add new leave request
   */
  public async addLeaveRequest(
    leaveRequest: ILeaveRequest
  ): Promise<void> {

    try {

      await this.sp.web.lists
        .getByTitle("LeaveRequest")
        .items
        .add({

          Title: leaveRequest.Title,

          Employee: leaveRequest.Employee,

          LeaveType: leaveRequest.LeaveType,

          StartDate: leaveRequest.StartDate,

          EndDate: leaveRequest.EndDate,

          Reason: leaveRequest.Reason,

          Status: "Pending"

        });

    } catch (error) {

      console.error(
        "Error while adding leave request:",
        error
      );

      throw error;
    }
  }


  /**
   * Update leave request
   */
  public async updateLeaveRequest(
    id: number,
    leaveRequest: ILeaveRequest
  ): Promise<void> {

    try {

      await this.sp.web.lists
        .getByTitle("LeaveRequest")
        .items
        .getById(id)
        .update({

          Title: leaveRequest.Title,

          Employee: leaveRequest.Employee,

          LeaveType: leaveRequest.LeaveType,

          StartDate: leaveRequest.StartDate,

          EndDate: leaveRequest.EndDate,

          Reason: leaveRequest.Reason

        });

    } catch (error) {

      console.error(
        "Error while updating leave request:",
        error
      );

      throw error;
    }
  }


  /**
   * Delete leave request
   */
  public async deleteLeaveRequest(
    id: number
  ): Promise<void> {

    try {

      await this.sp.web.lists
        .getByTitle("LeaveRequest")
        .items
        .getById(id)
        .delete();

    } catch (error) {

      console.error(
        "Error while deleting leave request:",
        error
      );

      throw error;
    }
  }

}