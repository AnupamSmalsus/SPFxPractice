import React from "react";
import { useDispatch, useSelector } from "react-redux";
import { updateField, resetForm } from "../state/students/studentsSlice";
import { AppDispatch } from "../state/store";
import { Web } from "sp-pnp-js";
import {
  Stack,
  TextField,
  DatePicker,
  PrimaryButton,
  DefaultButton
} from "@fluentui/react";
import '../components/ReduxUserForm.module.scss'

const StudentForm = (props: any) => {
  const baseUrl = props?.siteUrl;
  const dispatch = useDispatch<AppDispatch>();

  const formData = useSelector(
    (state: any) => state.student
  );

  const handleSubmit = async (e: any) => {
    e.preventDefault();
    const web = new Web(baseUrl);
    await web.lists.getById(props?.TestUsersListId).items.add(formData).then((res: any) => {
      console.log("Item is added successfully", formData);
      dispatch(resetForm());
      props?.closePopupCallBack();
    }).catch((error: any) => {
      console.log(error)
    })
  };

  return (
    <Stack tokens={{ childrenGap: 20 }} className="formContainer">

      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <TextField
          label="First Name"
          value={formData.First_x0020_Name}
          onChange={(e, value) =>
            dispatch(updateField({
              field: "First_x0020_Name",
              value: value || ""
            }))
          }
          styles={{ root: { width: "100%" } }}
        />

        <TextField
          label="Last Name"
          value={formData.Last_x0020_Name}
          onChange={(e, value) =>
            dispatch(updateField({
              field: "Last_x0020_Name",
              value: value || ""
            }))
          }
          styles={{ root: { width: "100%" } }}
        />
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <TextField
          label="Student ID"
          value={formData.Title}
          onChange={(e, value) =>
            dispatch(updateField({
              field: "Title",
              value: value || ""
            }))
          }
          styles={{ root: { width: "100%" } }}
        />

        <TextField
          label="Class"
          value={formData.Class}
          onChange={(e, value) =>
            dispatch(updateField({
              field: "Class",
              value: value || ""
            }))
          }
          styles={{ root: { width: "100%" } }}
        />
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <TextField
          label="Section"
          value={formData.Section}
          onChange={(e, value) =>
            dispatch(updateField({
              field: "Section",
              value: value || ""
            }))
          }
          styles={{ root: { width: "100%" } }}
        />

        <DatePicker
          label="Date of Birth"
          value={
            formData.DateOfBirth
              ? new Date(formData.DateOfBirth)
              : undefined
          }
          onSelectDate={(date) =>
            dispatch(
              updateField({
                field: "DateOfBirth",
                value: date
                  ? date.toLocaleDateString("en-US")
                  : ""
              })
            )
          }
          styles={{ root: { width: "100%" } }}
        />
      </Stack>

      <TextField
        label="Address"
        multiline
        rows={4}
        onChange={(e, value) =>
          dispatch(updateField({
            field: "Address",
            value: value || ""
          }))
        }
        value={formData.Address}
      />

      <Stack
        horizontal
        horizontalAlign="end"
        tokens={{ childrenGap: 10 }}
      >
        <DefaultButton
          text="Reset"
          onClick={() => dispatch(resetForm())}
        />

        <PrimaryButton
          text="Save Student"
          onClick={handleSubmit}
        />
      </Stack>

    </Stack>
  );
};

export default StudentForm;