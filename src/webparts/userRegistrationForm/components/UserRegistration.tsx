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
    DefaultButton
} from '@fluentui/react';

const classNames = mergeStyleSets({
    topBar: {
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginBottom: '15px'
    }
});

const UserRegistration = ({ allProps }: any) => {
    const [userDetails, setUserDetails] = React.useState<any[]>([]);
    const [departmentChoices, setDepartmentChoices] = React.useState<string[]>([]);
    const [designationChoices, setDesignationChoices] = React.useState<string[]>([]);

    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [editUserId, setEditUserId] = React.useState<number | null>(null);
    const [formData, setFormData] = React.useState<any>({
        Title: '',
        Department: '',
        Designation: '',
        Address: '',
        JoiningDate: '',
        LeavingDate: ''
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

    const fetchUserDetails = async () => {
        const web = new Web(allProps?.siteUrl);
        try {
            const items: any = await web.lists.getById(allProps?.UserRegistrationDetailsList).items
                .select('Id', 'Title', 'Department', 'Designation', 'Address', 'JoiningDate', 'LeavingDate')
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
                LeavingDate: user.LeavingDate ? user.LeavingDate.split('T')[0] : ''
            });
        } else {
            setEditUserId(null);
            setFormData({
                Title: '',
                Department: '',
                Designation: '',
                Address: '',
                JoiningDate: '',
                LeavingDate: ''
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
                await list.items.add(userData);
            }
            setIsPanelOpen(false);
            fetchUserDetails();
        }
        catch (error) {
            console.error('Error submitting user details:', error);
        }
    };

    const deleteUser = async (user: any) => {
        const web = new Web(allProps?.siteUrl)
        await web.lists.getById(allProps?.UserRegistrationDetailsList).items.getById(user.Id).recycle();
        fetchUserDetails()
    }

    const columns: IColumn[] = [
        { key: 'Title', name: 'Name', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
        { key: 'Department', name: 'Department', fieldName: 'Department', minWidth: 100, maxWidth: 150, isResizable: true },
        { key: 'Designation', name: 'Designation', fieldName: 'Designation', minWidth: 100, maxWidth: 150, isResizable: true },
        { key: 'Address', name: 'Address', fieldName: 'Address', minWidth: 100, maxWidth: 200, isResizable: true },
        {
            key: 'JoiningDate',
            name: 'Joining Date',
            fieldName: 'JoiningDate',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            onRender: (item) => <span>{item.JoiningDate ? new Date(item.JoiningDate).toLocaleDateString() : ''}</span>
        },
        {
            key: 'LeavingDate',
            name: 'Leaving Date',
            fieldName: 'LeavingDate',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            onRender: (item) => <span>{item.LeavingDate ? new Date(item.LeavingDate).toLocaleDateString() : ''}</span>
        },
        {
            key: 'Action',
            name: 'Action',
            fieldName: 'Action',
            minWidth: 50,
            maxWidth: 50,
            onRender: (item) => (
                <>
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
                </>
            )
        }
    ];

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
        <div>
            <div className={classNames.topBar}>
                <h2>Registered Users</h2>
                <PrimaryButton text="Add User" iconProps={{ iconName: 'Add' }} onClick={() => openPanel()} />
            </div>

            <DetailsList
                items={userDetails}
                columns={columns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
            />

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