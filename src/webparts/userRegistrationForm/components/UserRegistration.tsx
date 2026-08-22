import * as React from 'react';
import { Web } from 'sp-pnp-js';
import {
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    PrimaryButton,
    IconButton,
    Panel,
    PanelType,
    TextField,
    Dropdown,
    Stack,
    IStackTokens,
    mergeStyleSets,
    SelectionMode,
    DefaultButton,
    CommandBar,
    ICommandBarItemProps,
    SearchBox,
    Text,
    getTheme,
    FontWeights,
    MessageBar,
    MessageBarType
} from '@fluentui/react';

const theme = getTheme();
const classNames = mergeStyleSets({
    container: {
        width: '100%',
        maxWidth: '100%',
        boxSizing: 'border-box',
        padding: '24px',
        backgroundColor: theme.palette.white,
        boxShadow: theme.effects.elevation16,
        borderRadius: theme.effects.roundedCorner6,
        margin: '16px 0'
    },
    detailsListWrapper: {
        width: '100%',
        maxWidth: '100%',
    },
    detailsList: {
        width: '100%',
        maxWidth: '100%',
        selectors: {
            '.ms-DetailsHeader': {
                width: '100% !important',
            },
            '.ms-DetailsRow-fields': {
                width: '100% !important',
                minWidth: '100% !important',
            },
        },
    },
    headerStack: {
        marginBottom: '20px'
    },
    statusBadgeApproved: {
        backgroundColor: '#dff6dd',
        color: '#107c10',
        padding: '4px 10px',
        borderRadius: '12px',
        fontSize: '12px',
        fontWeight: FontWeights.semibold,
        display: 'inline-block'
    },
    statusBadgePending: {
        backgroundColor: '#fff4ce',
        color: '#797775',
        padding: '4px 10px',
        borderRadius: '12px',
        fontSize: '12px',
        fontWeight: FontWeights.semibold,
        display: 'inline-block'
    },
    searchBox: {
        width: '280px',
    }
});

const ACTION_COLUMN_WIDTH = 90;
const DATA_COLUMN_COUNT = 6;
const MIN_DATA_COLUMN_WIDTH = 100;

const UserRegistration = ({ allProps }: any) => {
    const [userDetails, setUserDetails] = React.useState<any[]>([]);
    const [departmentChoices, setDepartmentChoices] = React.useState<string[]>([]);
    const [designationChoices, setDesignationChoices] = React.useState<string[]>([]);
    const [searchTerm, setSearchTerm] = React.useState('');
    const listWrapperRef = React.useRef<HTMLDivElement>(null);
    const [listWidth, setListWidth] = React.useState(0);

    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [editUserId, setEditUserId] = React.useState<number | null>(null);
    const [formData, setFormData] = React.useState<any>({
        Title: '',
        Department: '',
        Designation: '',
        Address: '',
        JoiningDate: '',
        LeavingDate: '',
        Status: 'Pending',
        ApproverEmail: 'anupamrawat17@gmail.com'
    });

    React.useEffect(() => {
        fetchUserDetails();
        const init = async () => {
            const depChoices = await getChoicesFromField(allProps?.UserRegistrationDetailsList, 'Department');
            setDepartmentChoices(depChoices || []);
            const desChoices = await getChoicesFromField(allProps?.UserRegistrationDetailsList, 'Designation');
            setDesignationChoices(desChoices || []);
        }
        init();
    }, []);

    React.useEffect(() => {
        const el = listWrapperRef.current;
        if (!el) {
            return;
        }

        const updateWidth = (): void => {
            setListWidth(el.offsetWidth);
        };

        updateWidth();
        const observer = new ResizeObserver(updateWidth);
        observer.observe(el);
        return () => observer.disconnect();
    }, []);

    const fetchUserDetails = async () => {
        const web = new Web(allProps?.siteUrl);
        try {
            const items: any = await web.lists.getById(allProps?.UserRegistrationDetailsList).items
                .select('Id', 'Title', 'Department', 'Designation', 'Address', 'JoiningDate', 'LeavingDate', 'Status', 'ApproverEmail')
                .filter("Status eq 'Approved' or Status eq 'Pending'")
                .getAll();
            setUserDetails(items);
        }
        catch (error) {
            console.error('Error fetching user details:', error);
        }
    }

    const getChoicesFromField = async (listId: any, columnName: any) => {
        const web = new Web(allProps?.siteUrl);
        try {
            const field = await web.lists.getById(listId).fields.getByInternalNameOrTitle(columnName).get();
            return field.Choices;
        } catch (error) {
            console.error('Error getting choices:', error);
            return [];
        }
    };

    const openPanel = (user?: any) => {
        if (user) {
            setEditUserId(user.Id);
            setFormData({
                Title: user.Title || '',
                Department: user.Department || '',
                Designation: user.Designation || '',
                Address: user.Address || '',
                JoiningDate: user.JoiningDate ? user.JoiningDate.split('T')[0] : '',
                LeavingDate: user.LeavingDate ? user.LeavingDate.split('T')[0] : '',
                Status: user.Status || 'Pending',
                ApproverEmail: user.ApproverEmail || 'anupamrawat17@gmail.com'
            });
        } else {
            setEditUserId(null);
            setFormData({
                Title: '',
                Department: '',
                Designation: '',
                Address: '',
                JoiningDate: '',
                LeavingDate: '',
                Status: 'Pending',
                ApproverEmail: 'anupamrawat17@gmail.com'
            });
        }
        setIsPanelOpen(true);
    };

    const closePanel = () => {
        setIsPanelOpen(false);
    };

    const handleSubmit = async () => {
        if (!formData.Title || !formData.Department || !formData.Designation || !formData.Address || !formData.JoiningDate) {
            alert('Please fill out all required fields.');
            return;
        }

        const userData: any = { ...formData };
        if (!userData.LeavingDate) {
            userData.LeavingDate = null;
        }

        try {
            const web = new Web(allProps?.siteUrl);
            const list = web.lists.getById(allProps?.UserRegistrationDetailsList);
            if (editUserId) {
                await list.items.getById(editUserId).update(userData);
            } else {
                await list.items.add(userData)
            }
            setIsPanelOpen(false);
            fetchUserDetails();
        }
        catch (error) {
            console.error('Error submitting user details:', error);
        }
    };

    const deleteUser = async (user: any) => {
        const confirm = window.confirm(`Are you sure you want to delete the user ${user.Title}?`);
        if (!confirm) {
            return;
        }
        const web = new Web(allProps?.siteUrl)
        await web.lists.getById(allProps?.UserRegistrationDetailsList).items.getById(user.Id).recycle();
        fetchUserDetails()
    }

    const filteredUsers = React.useMemo(() => {
        if (!searchTerm) return userDetails;
        const lowerSearch = searchTerm.toLowerCase();
        return userDetails.filter(u =>
            (u.Title && u.Title.toLowerCase().includes(lowerSearch)) ||
            (u.Department && u.Department.toLowerCase().includes(lowerSearch)) ||
            (u.Designation && u.Designation.toLowerCase().includes(lowerSearch))
        );
    }, [userDetails, searchTerm]);

    const commandBarItems: ICommandBarItemProps[] = [
        {
            key: 'newItem',
            text: 'Add User',
            iconProps: { iconName: 'Add' },
            onClick: () => openPanel(),
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: () => fetchUserDetails(),
        }
    ];

    const farCommandBarItems: ICommandBarItemProps[] = [
        {
            key: 'search',
            onRender: () => (
                <SearchBox
                    placeholder="Search users..."
                    className={classNames.searchBox}
                    onChange={(_, newValue) => setSearchTerm(newValue || '')}
                    styles={{ root: { marginTop: 4 } }}
                />
            )
        }
    ];

    const dataColumnWidth = React.useMemo(() => {
        if (listWidth <= ACTION_COLUMN_WIDTH) {
            return MIN_DATA_COLUMN_WIDTH;
        }
        return Math.max(
            MIN_DATA_COLUMN_WIDTH,
            Math.floor((listWidth - ACTION_COLUMN_WIDTH) / DATA_COLUMN_COUNT)
        );
    }, [listWidth]);

    const columns: IColumn[] = React.useMemo(() => [
        { key: 'Title', name: 'Name', fieldName: 'Title', minWidth: dataColumnWidth, maxWidth: dataColumnWidth, isResizable: false },
        { key: 'Department', name: 'Department', fieldName: 'Department', minWidth: dataColumnWidth, maxWidth: dataColumnWidth, isResizable: false },
        { key: 'Designation', name: 'Designation', fieldName: 'Designation', minWidth: dataColumnWidth, maxWidth: dataColumnWidth, isResizable: false },
        {
            key: 'JoiningDate',
            name: 'Joining Date',
            fieldName: 'JoiningDate',
            minWidth: dataColumnWidth,
            maxWidth: dataColumnWidth,
            isResizable: false,
            onRender: (item) => <span>{item.JoiningDate ? new Date(item.JoiningDate).toLocaleDateString() : ''}</span>
        },
        {
            key: 'LeavingDate',
            name: 'Leaving Date',
            fieldName: 'LeavingDate',
            minWidth: dataColumnWidth,
            maxWidth: dataColumnWidth,
            isResizable: false,
            onRender: (item) => <span>{item.LeavingDate ? new Date(item.LeavingDate).toLocaleDateString() : ''}</span>
        },
        {
            key: 'Status',
            name: 'Status',
            fieldName: 'Status',
            minWidth: dataColumnWidth,
            maxWidth: dataColumnWidth,
            isResizable: false,
            onRender: (item) => (
                <span className={item.Status === 'Approved' ? classNames.statusBadgeApproved : classNames.statusBadgePending}>
                    {item.Status}
                </span>
            )
        },
        {
            key: 'Action',
            name: 'Action',
            fieldName: 'Action',
            minWidth: ACTION_COLUMN_WIDTH,
            maxWidth: ACTION_COLUMN_WIDTH,
            isResizable: false,
            onRender: (item) => (
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                    <IconButton
                        iconProps={{ iconName: 'Edit' }}
                        title="Edit"
                        ariaLabel="Edit"
                        onClick={() => openPanel(item)}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Delete"
                        ariaLabel="Delete"
                        onClick={() => deleteUser(item)}
                    />
                </Stack>
            )
        }
    ], [dataColumnWidth]);

    const stackTokens: IStackTokens = { childrenGap: 15 };

    const onRenderFooterContent = React.useCallback(
        () => (
            <div>
                <PrimaryButton onClick={handleSubmit} style={{ marginRight: 8 }}>
                    {editUserId ? 'Update' : 'Save'}
                </PrimaryButton>
                <DefaultButton onClick={closePanel}>Cancel</DefaultButton>
            </div>
        ),
        [handleSubmit, closePanel, editUserId]
    );

    return (
        <div className={classNames.container}>
            <Stack className={classNames.headerStack} tokens={{ childrenGap: 5 }}>
                <Text variant="xLarge" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                    User Registration
                </Text>
                <Text variant="medium" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                    Manage all registered users, their roles, and system statuses.
                </Text>
            </Stack>

            <CommandBar
                items={commandBarItems}
                farItems={farCommandBarItems}
                ariaLabel="User management actions"
                styles={{ root: { padding: 0, marginBottom: 16, width: '100%' } }}
            />

            <div ref={listWrapperRef} className={classNames.detailsListWrapper}>
                {filteredUsers.length === 0 ? (
                    <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
                        No users found matching the criteria.
                    </MessageBar>
                ) : (
                    <DetailsList
                        className={classNames.detailsList}
                        items={filteredUsers}
                        columns={columns}
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        selectionMode={SelectionMode.none}
                        styles={{
                            root: { width: '100%', maxWidth: '100%' },
                            focusZone: { width: '100%', maxWidth: '100%' },
                            contentWrapper: { width: '100%', maxWidth: '100%' },
                        }}
                    />
                )}
            </div>

            <Panel
                isOpen={isPanelOpen}
                onDismiss={closePanel}
                headerText={editUserId ? "Edit User" : "Add User"}
                type={PanelType.medium}
                onRenderFooterContent={onRenderFooterContent}
                isFooterAtBottom={true}
            >
                <Stack tokens={stackTokens}>
                    <TextField
                        label="Name"
                        required
                        value={formData.Title}
                        onChange={(e, newValue) => setFormData({ ...formData, Title: newValue || '' })}
                    />
                    <Dropdown
                        label="Department"
                        required
                        options={departmentChoices.map(d => ({ key: d, text: d }))}
                        selectedKey={formData.Department}
                        onChange={(e, option) => setFormData({ ...formData, Department: option?.key as string || '' })}
                    />
                    <Dropdown
                        label="Designation"
                        required
                        options={designationChoices.map(d => ({ key: d, text: d }))}
                        selectedKey={formData.Designation}
                        onChange={(e, option) => setFormData({ ...formData, Designation: option?.key as string || '' })}
                    />
                    <TextField
                        label="Address"
                        required
                        multiline
                        rows={3}
                        value={formData.Address}
                        onChange={(e, newValue) => setFormData({ ...formData, Address: newValue || '' })}
                    />
                    <TextField
                        label="Joining Date"
                        type="date"
                        required
                        value={formData.JoiningDate}
                        onChange={(e, newValue) => setFormData({ ...formData, JoiningDate: newValue || '' })}
                    />
                    <TextField
                        label="Leaving Date"
                        type="date"
                        value={formData.LeavingDate}
                        onChange={(e, newValue) => setFormData({ ...formData, LeavingDate: newValue || '' })}
                    />
                </Stack>
            </Panel>
        </div>
    );
}

export default UserRegistration;