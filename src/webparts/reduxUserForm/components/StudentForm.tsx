import React from "react";
import { useDispatch, useSelector } from "react-redux";
import { updateField, resetForm } from "../state/students/studentsSlice";
import { AppDispatch } from "../state/store";
import { Web } from "sp-pnp-js";

const StudentForm = (props: any) => {
  const baseUrl = props?.props?.siteUrl;
  const dispatch = useDispatch<AppDispatch>();

  const formData = useSelector(
    (state: any) => state.student
  );

  const handleChange = (e: any) => {
    dispatch(
      updateField({
        field: e.target.name,
        value: e.target.value
      })
    );
  };

  const handleSubmit = async (e: any) => {
    e.preventDefault();
    const web = new Web(baseUrl);
    await web.lists.getById(props?.props?.TestUsersListId).items.add(formData).then((res: any) => {
        console.log("Item is added successfully")
        dispatch(resetForm());
    }).catch((error: any) => {
        console.log(error)
    })
    console.log(formData);
  };

  return (
    <form onSubmit={handleSubmit}>
      <input
        name="Title"
        placeholder="Title"
        value={formData.Title}
        onChange={handleChange}
      />

      <input
        name="First_x0020_Name"
        placeholder="First Name"
        value={formData.First_x0020_Name}
        onChange={handleChange}
      />

      <input
        name="Last_x0020_Name"
        placeholder="Last Name"
        value={formData.Last_x0020_Name}
        onChange={handleChange}
      />

      <input
        name="Class"
        placeholder="Class"
        value={formData.Class}
        onChange={handleChange}
      />

      <input
        name="Section"
        placeholder="Section"
        value={formData.Section}
        onChange={handleChange}
      />

      <input
        type="date"
        name="DateOfBirth"
        value={formData.DateOfBirth}
        onChange={handleChange}
      />

      <textarea
        name="Address"
        placeholder="Address"
        value={formData.Address}
        onChange={handleChange}
      />

      <button type="submit">
        Save
      </button>

      <button
        type="button"
        onClick={() => dispatch(resetForm())}
      >
        Reset
      </button>
    </form>
  );
};

export default StudentForm;