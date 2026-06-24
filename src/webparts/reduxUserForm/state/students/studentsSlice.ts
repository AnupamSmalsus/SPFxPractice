import { createSlice, PayloadAction } from '@reduxjs/toolkit'


export interface UserDetailsState {
    Title: string;
    Class: string;
    Section: string;
    First_x0020_Name: string;
    Last_x0020_Name: string;
    DateOfBirth: string;
    Address: string;
}
const initialState: UserDetailsState = {
    Title: "",
    First_x0020_Name: "",
    Last_x0020_Name: "",
    Class: "",
    Section: "",
    DateOfBirth: "",
    Address: ""
}

const studentSlice = createSlice({
    name: 'stundents',
    initialState,
    reducers: {
        updateField: (state, action: PayloadAction<{
            field: keyof UserDetailsState;
            value: string;
        }>) => {
            const { field, value } = action.payload;
            state[field] = value;
        },
        resetForm: () => initialState
    }
})

export const { updateField, resetForm } = studentSlice.actions;

export default studentSlice.reducer;
