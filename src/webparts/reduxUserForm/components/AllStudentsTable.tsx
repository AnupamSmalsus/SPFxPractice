import * as React from "react";
import { Web } from "sp-pnp-js";
import {
  PrimaryButton,
  Panel,
  PanelType,
  TextField
} from "@fluentui/react";
import StudentForm from "./StudentForm";

import {
  createColumnHelper,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  SortingState,
  useReactTable
} from "@tanstack/react-table";

interface IStudent {
  Id: number;
  Title: string;
  First_x0020_Name: string;
  Last_x0020_Name: string;
  FullName: string;
  Class: string;
  Section: string;
  DateOfBirth: string;
  Address: string;
}

const StudentsTable = (props: any) => {
  const [allTableData, setAllTableData] = React.useState<IStudent[]>([]);
  const [openPanel, setOpenPanel] = React.useState(false);
  const [globalFilter, setGlobalFilter] = React.useState("");
  const [sorting, setSorting] = React.useState<SortingState>([]);

  React.useEffect(() => {
    getAllUsersData();
  }, []);

  const getAllUsersData = async () => {
    try {
      const web = new Web(props?.props?.siteUrl);

      const res = await web.lists
        .getById(props?.props?.TestUsersListId)
        .items.select(
          "Id",
          "Title",
          "First_x0020_Name",
          "Last_x0020_Name",
          "Class",
          "Section",
          "DateOfBirth",
          "Address"
        )
        .getAll();

      const updatedTableData: IStudent[] = res.map((item: any) => ({
        ...item,
        FullName: `${item.First_x0020_Name || ""} ${
          item.Last_x0020_Name || ""
        }`.trim()
      }));

      setAllTableData(updatedTableData);
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  const columnHelper = createColumnHelper<IStudent>();

  const columns = React.useMemo(
    () => [

      columnHelper.accessor("Title", {
        header: "Student ID"
      }),

      columnHelper.accessor("First_x0020_Name", {
        header: "First Name"
      }),

      columnHelper.accessor("Last_x0020_Name", {
        header: "Last Name"
      }),

      columnHelper.accessor("FullName", {
        header: "Full Name"
      }),

      columnHelper.accessor("Class", {
        header: "Class"
      }),

      columnHelper.accessor("Section", {
        header: "Section"
      }),

      columnHelper.accessor("DateOfBirth", {
        header: "Date Of Birth",
        cell: info => {
          const value = info.getValue();

          return value
            ? new Date(value).toLocaleDateString("en-GB")
            : "";
        }
      }),

      columnHelper.accessor("Address", {
        header: "Address"
      })
    ],
    []
  );

  const table = useReactTable({
    data: allTableData,
    columns,

    state: {
      globalFilter,
      sorting
    },

    onGlobalFilterChange: setGlobalFilter,
    onSortingChange: setSorting,

    globalFilterFn: (row, columnId, filterValue) => {
      const search = String(filterValue).toLowerCase();

      return Object.values(row.original).some((value: any) =>
        String(value || "")
          .toLowerCase()
          .includes(search)
      );
    },

    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getSortedRowModel: getSortedRowModel()
  });

  const closeForm = () => {
    setOpenPanel(false);

    // Refresh table after add/edit
    getAllUsersData();
  };

  return (
    <>
      <h1>Welcome To Students Portal</h1>

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          marginBottom: "20px",
          alignItems: "end"
        }}
      >
        <div style={{ width: "350px" }}>
          <TextField
            label="Search Students"
            placeholder="Search by any field..."
            value={globalFilter}
            onChange={(_, value) =>
              setGlobalFilter(value || "")
            }
          />
        </div>

        <PrimaryButton
          text="Add Student"
          onClick={() => setOpenPanel(true)}
        />
      </div>

      <div style={{ overflowX: "auto" }}>
        <table
          style={{
            width: "100%",
            borderCollapse: "collapse",
            border: "1px solid #ddd"
          }}
        >
          <thead>
            {table.getHeaderGroups().map(headerGroup => (
              <tr key={headerGroup.id}>
                {headerGroup.headers.map(header => (
                  <th
                    key={header.id}
                    onClick={header.column.getToggleSortingHandler()}
                    style={{
                      border: "1px solid #ddd",
                      padding: "12px",
                      backgroundColor: "#f3f2f1",
                      cursor: "pointer",
                      textAlign: "left",
                      userSelect: "none"
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center"
                      }}
                    >
                      {flexRender(
                        header.column.columnDef.header,
                        header.getContext()
                      )}

                      {{
                        asc: "🔼",
                        desc: "🔽"
                      }[
                        header.column.getIsSorted() as string
                      ] ?? "↕️"}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>

          <tbody>
            {table.getRowModel().rows.length > 0 ? (
              table.getRowModel().rows.map(row => (
                <tr key={row.id}>
                  {row.getVisibleCells().map(cell => (
                    <td
                      key={cell.id}
                      style={{
                        border: "1px solid #ddd",
                        padding: "10px"
                      }}
                    >
                      {flexRender(
                        cell.column.columnDef.cell,
                        cell.getContext()
                      )}
                    </td>
                  ))}
                </tr>
              ))
            ) : (
              <tr>
                <td
                  colSpan={columns.length}
                  style={{
                    textAlign: "center",
                    padding: "20px"
                  }}
                >
                  No records found
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <Panel
        isOpen={openPanel}
        onDismiss={() => setOpenPanel(false)}
        headerText="Student Registration Form"
        type={PanelType.medium}
      >
        <StudentForm
          closePopupCallBack={closeForm}
          siteUrl={props?.props?.siteUrl}
          TestUsersListId={props?.props?.TestUsersListId}
        />
      </Panel>
    </>
  );
};

export default StudentsTable;