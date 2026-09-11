import * as React from "react";

import {
    useEffect,
    useState
} from "react";

import {
    PrimaryButton,
    DefaultButton,
    TextField,
    Dropdown,
    IDropdownOption,
    DatePicker,
    Panel,
    PanelType,
    DetailsList,
    IColumn,
    SelectionMode,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize,
    IconButton
} from "@fluentui/react";

import {
    ILeaveRequest
} from "./ILeaveRequest";

import {
    SharePointService
} from "../service/SharepointService";

import styles from "./LeaveRequestPortal.module.scss";

import { ILeaveRequestPortalProps } from "./ILeaveRequestPortalProps";


const LeaveRequestPortal: React.FC<ILeaveRequestPortalProps> = (
    props
) => {

    const [leaveRequests, setLeaveRequests] =
        useState<ILeaveRequest[]>([]);

    const [filteredRequests, setFilteredRequests] =
        useState<ILeaveRequest[]>([]);


    const [loading, setLoading] =
        useState<boolean>(true);


    const [saving, setSaving] =
        useState<boolean>(false);


    const [error, setError] =
        useState<string>("");


    const [success, setSuccess] =
        useState<string>("");


    const [isPanelOpen, setIsPanelOpen] =
        useState<boolean>(false);


    const [isEditMode, setIsEditMode] =
        useState<boolean>(false);


    const [selectedRequestId, setSelectedRequestId] =
        useState<number | undefined>(undefined);


    const [searchText, setSearchText] =
        useState<string>("");


    const [statusFilter, setStatusFilter] =
        useState<string>("All");


    const [employee, setEmployee] =
        useState<string>("");


    const [leaveType, setLeaveType] =
        useState<string>("");


    const [startDate, setStartDate] =
        useState<any>(undefined);


    const [endDate, setEndDate] =
        useState<any>(undefined);


    const [reason, setReason] =
        useState<string>("");


    const [formErrors, setFormErrors] =
        useState<string[]>([]);


    const service = React.useMemo(
        () => new SharePointService(props.context),
        [props.context]
    );


    const leaveTypeOptions: IDropdownOption[] = [

        {
            key: "Casual Leave",
            text: "Casual Leave"
        },

        {
            key: "Sick Leave",
            text: "Sick Leave"
        },

        {
            key: "Earned Leave",
            text: "Earned Leave"
        },

        {
            key: "Work From Home",
            text: "Work From Home"
        },

        {
            key: "Other",
            text: "Other"
        }

    ];


    const statusOptions: IDropdownOption[] = [

        {
            key: "All",
            text: "All"
        },

        {
            key: "Pending",
            text: "Pending"
        },

        {
            key: "Approved",
            text: "Approved"
        },

        {
            key: "Rejected",
            text: "Rejected"
        }

    ];


    /**
     * Load requests
     */
    const loadLeaveRequests = async (): Promise<void> => {

        try {

            setLoading(true);

            setError("");

            const data =
                await service.getLeaveRequests();

            setLeaveRequests(data);

            setFilteredRequests(data);

        } catch (err) {

            console.error(err);

            setError(
                "Unable to load leave requests."
            );

        } finally {

            setLoading(false);

        }
    };


    useEffect(() => {

        loadLeaveRequests();

    }, []);


    /**
     * Filter requests
     */
    useEffect(() => {

        let result = [...leaveRequests];


        if (searchText.trim()) {

            const search =
                searchText.toLowerCase();

            result = result.filter(
                (item) =>
                    item.Employee
                        .toLowerCase()
                        .includes(search) ||

                    item.LeaveType
                        .toLowerCase()
                        .includes(search) ||

                    item.Reason
                        .toLowerCase()
                        .includes(search)
            );
        }


        if (statusFilter !== "All") {

            result = result.filter(
                (item) =>
                    item.Status === statusFilter
            );

        }


        setFilteredRequests(result);

    }, [
        searchText,
        statusFilter,
        leaveRequests
    ]);


    /**
     * Reset form
     */
    const resetForm = (): void => {

        setEmployee("");

        setLeaveType("");

        setStartDate(undefined);

        setEndDate(undefined);

        setReason("");

        setSelectedRequestId(undefined);

        setIsEditMode(false);

        setFormErrors([]);

    };


    /**
     * Open Add Panel
     */
    const openAddPanel = (): void => {

        resetForm();

        setIsPanelOpen(true);

    };


    /**
     * Open Edit Panel
     */
    const openEditPanel = (
        request: ILeaveRequest
    ): void => {

        setSelectedRequestId(
            request.Id
        );

        setEmployee(
            request.Employee
        );

        setLeaveType(
            request.LeaveType
        );

        setStartDate(
            request.StartDate
                ? new Date(request.StartDate)
                : undefined
        );

        setEndDate(
            request.EndDate
                ? new Date(request.EndDate)
                : undefined
        );

        setReason(
            request.Reason
        );

        setIsEditMode(true);

        setFormErrors([]);

        setIsPanelOpen(true);

    };


    /**
     * Validate form
     */
    const validateForm = (): boolean => {

        const errors: string[] = [];


        if (!employee.trim()) {

            errors.push(
                "Employee name is required."
            );

        }


        if (!leaveType) {

            errors.push(
                "Leave type is required."
            );

        }


        if (!startDate) {

            errors.push(
                "Start date is required."
            );

        }


        if (!endDate) {

            errors.push(
                "End date is required."
            );

        }


        if (
            startDate &&
            endDate &&
            endDate < startDate
        ) {

            errors.push(
                "End date cannot be before start date."
            );

        }


        if (!reason.trim()) {

            errors.push(
                "Reason is required."
            );

        }


        setFormErrors(errors);


        return errors.length === 0;

    };


    /**
     * Save request
     */
    const saveRequest = async (): Promise<void> => {

        if (!validateForm()) {

            return;

        }


        try {

            setSaving(true);

            setError("");

            setSuccess("");


            const request: ILeaveRequest = {

                Title:
                    `${employee} - ${leaveType}`,

                Employee:
                    employee.trim(),

                LeaveType:
                    leaveType,

                StartDate:
                    startDate?.toISOString(),

                EndDate:
                    endDate?.toISOString(),

                Reason:
                    reason.trim(),

                Status:
                    "Pending"

            };


            if (
                isEditMode &&
                selectedRequestId
            ) {

                await service.updateLeaveRequest(
                    selectedRequestId,
                    request
                );

                setSuccess(
                    "Leave request updated successfully."
                );

            } else {

                await service.addLeaveRequest(
                    request
                );

                setSuccess(
                    "Leave request submitted successfully."
                );

            }


            setIsPanelOpen(false);

            resetForm();

            await loadLeaveRequests();

        } catch (err) {

            console.error(err);

            setError(
                "Unable to save leave request."
            );

        } finally {

            setSaving(false);

        }

    };


    /**
     * Delete request
     */
    const deleteRequest = async (
        request: ILeaveRequest
    ): Promise<void> => {

        if (!request.Id) {

            return;

        }


        const confirmed =
            window.confirm(
                "Are you sure you want to delete this leave request?"
            );


        if (!confirmed) {

            return;

        }


        try {

            setLoading(true);

            setError("");

            setSuccess("");


            await service.deleteLeaveRequest(
                request.Id
            );


            setSuccess(
                "Leave request deleted successfully."
            );


            await loadLeaveRequests();

        } catch (err) {

            console.error(err);

            setError(
                "Unable to delete leave request."
            );

        } finally {

            setLoading(false);

        }

    };


    /**
     * Table columns
     */
    const columns: IColumn[] = [

        {
            key: "employee",
            name: "Employee",
            fieldName: "Employee",
            minWidth: 120,
            maxWidth: 180,
            isResizable: true
        },

        {
            key: "leaveType",
            name: "Leave Type",
            fieldName: "LeaveType",
            minWidth: 120,
            maxWidth: 160,
            isResizable: true
        },

        {
            key: "startDate",
            name: "Start Date",
            fieldName: "StartDate",
            minWidth: 100,
            maxWidth: 120,
            onRender: (
                item: ILeaveRequest
            ) =>
                new Date(
                    item.StartDate
                ).toLocaleDateString()
        },

        {
            key: "endDate",
            name: "End Date",
            fieldName: "EndDate",
            minWidth: 100,
            maxWidth: 120,
            onRender: (
                item: ILeaveRequest
            ) =>
                new Date(
                    item.EndDate
                ).toLocaleDateString()
        },

        {
            key: "reason",
            name: "Reason",
            fieldName: "Reason",
            minWidth: 180,
            maxWidth: 300,
            isResizable: true
        },

        {
            key: "status",
            name: "Status",
            fieldName: "Status",
            minWidth: 100,
            maxWidth: 120,
            onRender: (
                item: ILeaveRequest
            ) => {

                let className =
                    styles.pending;

                if (
                    item.Status === "Approved"
                ) {

                    className =
                        styles.approved;

                }

                if (
                    item.Status === "Rejected"
                ) {

                    className =
                        styles.rejected;

                }

                return (

                    <span
                        className={className}
                    >
                        {item.Status}
                    </span>

                );

            }
        },

        {
            key: "actions",
            name: "Actions",
            minWidth: 100,
            maxWidth: 120,

            onRender: (
                item: ILeaveRequest
            ) => (

                <div
                    className={styles.actionButtons}
                >

                    {item.Status === "Pending" && (

                        <>

                            <IconButton
                                iconProps={{
                                    iconName: "Edit"
                                }}
                                title="Edit"
                                ariaLabel="Edit"
                                onClick={() =>
                                    openEditPanel(item)
                                }
                            />

                            <IconButton
                                iconProps={{
                                    iconName: "Delete"
                                }}
                                title="Delete"
                                ariaLabel="Delete"
                                onClick={() =>
                                    deleteRequest(item)
                                }
                            />

                        </>

                    )}

                </div>

            )
        }

    ];


    return (

        <div className={styles.leavePortal}>

            {/* Header */}

            <div className={styles.header}>

                <div>

                    <h1>
                        Leave Request Portal
                    </h1>

                    <p>
                        Submit and manage your leave requests
                    </p>

                </div>


                <PrimaryButton
                    text="New Leave Request"
                    iconProps={{
                        iconName: "Add"
                    }}
                    onClick={openAddPanel}
                />

            </div>


            {/* Messages */}

            {error && (

                <MessageBar
                    messageBarType={
                        MessageBarType.error
                    }
                    onDismiss={() =>
                        setError("")
                    }
                >
                    {error}
                </MessageBar>

            )}


            {success && (

                <MessageBar
                    messageBarType={
                        MessageBarType.success
                    }
                    onDismiss={() =>
                        setSuccess("")
                    }
                >
                    {success}
                </MessageBar>

            )}


            {/* Filters */}

            <div className={styles.filters}>

                <TextField
                    placeholder="Search employee, leave type or reason..."
                    value={searchText}
                    onChange={(_, value) =>
                        setSearchText(value || "")
                    }
                    styles={{
                        root: {
                            width: 300
                        }
                    }}
                />


                <Dropdown
                    placeholder="Filter by status"
                    selectedKey={statusFilter}
                    options={statusOptions}
                    onChange={(_, option) =>
                        setStatusFilter(
                            option?.key as string
                        )
                    }
                    styles={{
                        root: {
                            width: 180
                        }
                    }}
                />

            </div>


            {/* Content */}

            <div className={styles.content}>

                {loading ? (

                    <div
                        className={styles.loader}
                    >

                        <Spinner
                            size={SpinnerSize.medium}
                            label="Loading leave requests..."
                        />

                    </div>

                ) : (

                    <DetailsList

                        items={
                            filteredRequests
                        }

                        columns={columns}

                        selectionMode={
                            SelectionMode.none
                        }

                        compact={false}

                    />

                )}

                {!loading &&
                    filteredRequests.length === 0 && (

                        <div
                            className={styles.emptyState}
                        >

                            No leave requests found.

                        </div>

                    )}

            </div>


            {/* Add/Edit Panel */}

            <Panel

                isOpen={isPanelOpen}

                type={PanelType.medium}

                headerText={
                    isEditMode
                        ? "Edit Leave Request"
                        : "New Leave Request"
                }

                onDismiss={() => {

                    setIsPanelOpen(false);

                    resetForm();

                }}

                closeButtonAriaLabel="Close"

                isFooterAtBottom={true}

                onRenderFooterContent={() => (

                    <div
                        className={styles.panelFooter}
                    >

                        <PrimaryButton
                            text={
                                saving
                                    ? "Saving..."
                                    : "Submit Request"
                            }
                            disabled={saving}
                            onClick={saveRequest}
                        />

                        <DefaultButton
                            text="Cancel"
                            disabled={saving}
                            onClick={() => {

                                setIsPanelOpen(false);

                                resetForm();

                            }}
                        />

                    </div>

                )}

            >

                {/* Validation */}

                {formErrors.length > 0 && (

                    <MessageBar
                        messageBarType={
                            MessageBarType.error
                        }
                    >

                        <ul>

                            {formErrors.map(
                                (error, index) => (

                                    <li key={index}>
                                        {error}
                                    </li>

                                )
                            )}

                        </ul>

                    </MessageBar>

                )}


                <div
                    className={styles.form}
                >

                    <TextField
                        label="Employee"
                        required
                        placeholder="Enter employee name"
                        value={employee}
                        onChange={(_, value) =>
                            setEmployee(value || "")
                        }
                    />


                    <Dropdown
                        label="Leave Type"
                        required
                        placeholder="Select leave type"
                        selectedKey={leaveType}
                        options={leaveTypeOptions}
                        onChange={(_, option) =>
                            setLeaveType(
                                option?.key as string
                            )
                        }
                    />


                    <DatePicker
                        label="Start Date"
                        isRequired
                        placeholder="Select start date"
                        value={startDate}
                        onSelectDate={(date) =>
                            setStartDate(date || undefined)
                        }
                        formatDate={(date) =>
                            date
                                ? date.toLocaleDateString()
                                : ""
                        }
                    />


                    <DatePicker
                        label="End Date"
                        isRequired
                        placeholder="Select end date"
                        value={endDate}
                        onSelectDate={(date) =>
                            setEndDate(date || undefined)
                        }
                        formatDate={(date) =>
                            date
                                ? date.toLocaleDateString()
                                : ""
                        }
                    />


                    <TextField
                        label="Reason"
                        required
                        multiline
                        rows={5}
                        placeholder="Enter reason for leave"
                        value={reason}
                        onChange={(_, value) =>
                            setReason(value || "")
                        }
                    />

                </div>

            </Panel>

        </div>

    );

};


export default LeaveRequestPortal;