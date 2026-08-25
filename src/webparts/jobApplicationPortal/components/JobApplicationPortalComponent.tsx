import * as React from 'react';
import * as service from '../service/service';
import {
    Panel,
    PanelType,
    PrimaryButton,
    DefaultButton,
    TextField,
    Stack,
    MessageBar,
    MessageBarType
} from '@fluentui/react';
import {
    ColumnDef,
    flexRender,
    getCoreRowModel,
    useReactTable
} from '@tanstack/react-table';

interface IJob {
    Id: number;
    Title: string;
    Description: string;
    Package: string;
    Experience: string;
    Vacancy: number;
}

interface IJobForm {
    Title: string;
    Description: string;
    Package: string;
    Experience: string;
    Vacancy: string;
}

interface IApplicationForm {
    ApplicantName: string;
    Email: string;
    Phone: string;
    Experience: string;
    CoverLetter: string;
    ExpectedSalary?: string;
}

const JobApplicationPortalComponent = (props: any) => {
    const [jobs, setJobs] = React.useState<IJob[]>([]);
    const [isJobPanelOpen, setIsJobPanelOpen] = React.useState(false);
    const [jobPanelMode, setJobPanelMode] = React.useState<'add' | 'edit'>('add');
    const [selectedJob, setSelectedJob] = React.useState<IJob | null>(null);
    const [isApplyPanelOpen, setIsApplyPanelOpen] = React.useState(false);
    const [isSaving, setIsSaving] = React.useState(false);
    const [isApplying, setIsApplying] = React.useState(false);
    const [message, setMessage] = React.useState('');
    const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.success);
    const [jobForm, setJobForm] = React.useState<IJobForm>({
        Title: '', Description: '', Package: '', Experience: '', Vacancy: ''
    });
    const [applicationForm, setApplicationForm] = React.useState<IApplicationForm>({
        ApplicantName: '', Email: '', Phone: '', Experience: '', CoverLetter: '', ExpectedSalary: ''
    });

    const fetchJobs = async () => {
        try {
            const data = await service.getJobs(props?.props);
            setJobs(data);
        } catch (error) {
            console.error('Error loading jobs:', error);
            setMessage('Unable to load jobs.');
            setMessageType(MessageBarType.error);
        }
    };

    React.useEffect(() => { fetchJobs(); }, []);

    const openAddJobPanel = () => {
        setJobPanelMode('add');
        setSelectedJob(null);
        setJobForm({ Title: '', Description: '', Package: '', Experience: '', Vacancy: '' });
        setMessage('');
        setIsJobPanelOpen(true);
    };

    const openEditJobPanel = (job: IJob) => {
        setJobPanelMode('edit');
        setSelectedJob(job);
        setJobForm({
            Title: job.Title || '',
            Description: job.Description || '',
            Package: job.Package || '',
            Experience: job.Experience || '',
            Vacancy: job.Vacancy?.toString() || ''
        });
        setMessage('');
        setIsJobPanelOpen(true);
    };

    const updateJobForm = (field: keyof IJobForm, value: string) => {
        setJobForm(prev => ({ ...prev, [field]: value }));
    };

    const saveJob = async () => {
        if (!jobForm.Title.trim()) return showError('Please enter job title.');
        if (!jobForm.Description.trim()) return showError('Please enter job description.');
        if (!jobForm.Package.trim()) return showError('Please enter salary package.');
        if (!jobForm.Experience.trim()) return showError('Please enter required experience.');
        if (!jobForm.Vacancy || Number(jobForm.Vacancy) <= 0) return showError('Please enter a valid number of vacancies.');

        try {
            setIsSaving(true);
            const jobData = {
                Title: jobForm.Title,
                Description: jobForm.Description,
                Package: jobForm.Package,
                Experience: jobForm.Experience,
                Vacancy: Number(jobForm.Vacancy)
            };

            if (jobPanelMode === 'add') {
                await service.addJob(props?.props, jobData);
                showSuccess('Job added successfully.');
            } else if (selectedJob) {
                await service.updateJob(props?.props, selectedJob.Id, jobData);
                showSuccess('Job updated successfully.');
            }

            await fetchJobs();
            setIsJobPanelOpen(false);
        } catch (error) {
            console.error('Error saving job:', error);
            showError(jobPanelMode === 'add' ? 'Unable to add job.' : 'Unable to update job.');
        } finally {
            setIsSaving(false);
        }
    };

    const deleteJob = async (job: IJob) => {
        if (!window.confirm(`Are you sure you want to delete "${job.Title}"?`)) return;
        try {
            await service.deleteJob(props?.props, job.Id);
            setJobs(prev => prev.filter(x => x.Id !== job.Id));
            showSuccess('Job deleted successfully.');
        } catch (error) {
            console.error('Error deleting job:', error);
            showError('Unable to delete job.');
        }
    };

    const openApplyPanel = (job: IJob) => {
        setSelectedJob(job);
        setApplicationForm({
            ApplicantName: '', Email: '', Phone: '', Experience: '', CoverLetter: '', ExpectedSalary: '',
        });
        setMessage('');
        setIsApplyPanelOpen(true);
    };

    const updateApplicationForm = (field: keyof IApplicationForm, value: string) => {
        setApplicationForm(prev => ({ ...prev, [field]: value }));
    };

    const submitApplication = async () => {
        if (!selectedJob) return;
        if (!applicationForm.ApplicantName.trim()) return showError('Please enter applicant name.');
        if (!applicationForm.Email.trim()) return showError('Please enter email.');
        if (!applicationForm.Phone.trim()) return showError('Please enter phone number.');
        if (!applicationForm.Experience.trim()) return showError('Please enter your experience.');
        if (!applicationForm.ExpectedSalary?.trim()) return showError('Please enter expected salary.');

        try {
            setIsApplying(true);
            await service.addApplication(props?.props, {
                JobID: selectedJob.Id,
                JobTitle: selectedJob.Title,
                ApplicantName: applicationForm.ApplicantName,
                Email: applicationForm.Email,
                Phone: applicationForm.Phone,
                Experience: applicationForm.Experience,
                ExpectedSalary: applicationForm.ExpectedSalary,
                CoverLetter: applicationForm.CoverLetter,
                AppliedDate: new Date().toISOString(),
                Status: 'Applied'
            });
            setIsApplyPanelOpen(false);
            showSuccess('Application submitted successfully.');
        } catch (error) {
            console.error('Error submitting application:', error);
            showError('Unable to submit application.');
        } finally {
            setIsApplying(false);
        }
    };

    const showError = (text: string) => {
        setMessage(text);
        setMessageType(MessageBarType.error);
    };

    const showSuccess = (text: string) => {
        setMessage(text);
        setMessageType(MessageBarType.success);
    };

    const columns = React.useMemo<ColumnDef<IJob>[]>(() => [
        {
            accessorKey: 'Title',
            header: 'Job Title',
            cell: info => <strong>{info.getValue<string>()}</strong>
        },
        {
            accessorKey: 'Description',
            header: 'Job Description',
            cell: info => (
                <div
                    title={info.getValue<string>()}
                    style={{
                        maxWidth: 450,
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap'
                    }}
                >
                    {info.getValue<string>()}
                </div>
            )
        },
        {
            accessorKey: 'Package',
            header: 'Salary Package',
            cell: info => info.getValue<string>()
        },
        {
            accessorKey: 'Experience',
            header: 'Experience',
            cell: info => info.getValue<string>()
        },
        {
            accessorKey: 'Vacancy',
            header: 'Vacancy',
            cell: info => info.getValue<number>()
        },
        {
            id: 'actions',
            header: 'Actions',
            cell: ({ row }) => (
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                    <PrimaryButton text="Apply" onClick={() => openApplyPanel(row.original)} />
                    <DefaultButton text="Edit" onClick={() => openEditJobPanel(row.original)} />
                    <DefaultButton text="Delete" onClick={() => deleteJob(row.original)} />
                </Stack>
            )
        }
    ], []);

    const table = useReactTable({
        data: jobs,
        columns,
        getCoreRowModel: getCoreRowModel()
    });

    return (
        <div style={{ width: '100%', padding: 24, boxSizing: 'border-box' }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 20 } }}>
                <h1 style={{ margin: 0 }}>Job Application Portal</h1>
                <PrimaryButton text="Add New Job" iconProps={{ iconName: 'Add' }} onClick={openAddJobPanel} />
            </Stack>

            {message && (
                <MessageBar
                    messageBarType={messageType}
                    onDismiss={() => setMessage('')}
                    styles={{ root: { marginBottom: 20 } }}
                >
                    {message}
                </MessageBar>
            )}

            <div style={{ width: '100%', overflowX: 'auto', border: '1px solid #edebe9' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                        {table.getHeaderGroups().map(headerGroup => (
                            <tr key={headerGroup.id}>
                                {headerGroup.headers.map(header => (
                                    <th key={header.id} style={headerStyle}>
                                        {header.isPlaceholder ? null : flexRender(header.column.columnDef.header, header.getContext())}
                                    </th>
                                ))}
                            </tr>
                        ))}
                    </thead>
                    <tbody>
                        {table.getRowModel().rows.map(row => (
                            <tr key={row.id} style={{ borderBottom: '1px solid #edebe9' }}>
                                {row.getVisibleCells().map(cell => (
                                    <td key={cell.id} style={cellStyle}>
                                        {flexRender(cell.column.columnDef.cell, cell.getContext())}
                                    </td>
                                ))}
                            </tr>
                        ))}
                        {!jobs.length && (
                            <tr>
                                <td colSpan={6} style={{ textAlign: 'center', padding: 30, color: '#605e5c' }}>
                                    No jobs available.
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>

            <Panel
                isOpen={isJobPanelOpen}
                onDismiss={() => !isSaving && setIsJobPanelOpen(false)}
                type={PanelType.medium}
                headerText={jobPanelMode === 'add' ? 'Add New Job' : 'Edit Job'}
                closeButtonAriaLabel="Close"
                isBlocking
            >
                <Stack tokens={{ childrenGap: 18 }}>
                    <TextField
                        label="Job Title"
                        required
                        value={jobForm.Title}
                        onChange={(_, value) => updateJobForm('Title', value || '')}
                    />
                    <TextField
                        label="Job Description"
                        required
                        multiline
                        rows={6}
                        value={jobForm.Description}
                        onChange={(_, value) => updateJobForm('Description', value || '')}
                    />
                    <TextField
                        label="Salary Package"
                        required
                        placeholder="Example: ₹8 - ₹12 LPA"
                        value={jobForm.Package}
                        onChange={(_, value) => updateJobForm('Package', value || '')}
                    />
                    <TextField
                        label="Experience"
                        required
                        placeholder="Example: 3+ Years"
                        value={jobForm.Experience}
                        onChange={(_, value) => updateJobForm('Experience', value || '')}
                    />
                    <TextField
                        label="Vacancies"
                        required
                        type="number"
                        min={1}
                        value={jobForm.Vacancy}
                        onChange={(_, value) => updateJobForm('Vacancy', value || '')}
                    />
                    <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 20 } }}>
                        <PrimaryButton
                            text={isSaving ? 'Saving...' : jobPanelMode === 'add' ? 'Add Job' : 'Update Job'}
                            disabled={isSaving}
                            onClick={saveJob}
                        />
                        <DefaultButton
                            text="Cancel"
                            disabled={isSaving}
                            onClick={() => setIsJobPanelOpen(false)}
                        />
                    </Stack>
                </Stack>
            </Panel>

            <Panel
                isOpen={isApplyPanelOpen}
                onDismiss={() => !isApplying && setIsApplyPanelOpen(false)}
                type={PanelType.medium}
                headerText="Apply for Job"
                closeButtonAriaLabel="Close"
                isBlocking
            >
                {selectedJob && (
                    <>
                    <Stack tokens={{ childrenGap: 18 }}>
                        <div style={jobInfoStyle}>
                            <h2 style={{ margin: '0 0 12px' }}>{selectedJob.Title}</h2>
                            <p><strong>Salary:</strong> {selectedJob.Package}</p>
                            <p><strong>Required Experience:</strong> {selectedJob.Experience}</p>
                            <p style={{ marginBottom: 0 }}><strong>Vacancies:</strong> {selectedJob.Vacancy}</p>
                        </div>

                        <TextField
                            label="Applicant Name"
                            required
                            value={applicationForm.ApplicantName}
                            onChange={(_, value) => updateApplicationForm('ApplicantName', value || '')}
                        />
                        <TextField
                            label="Email"
                            required
                            type="email"
                            value={applicationForm.Email}
                            onChange={(_, value) => updateApplicationForm('Email', value || '')}
                        />
                        <TextField
                            label="Phone Number"
                            required
                            value={applicationForm.Phone}
                            onChange={(_, value) => updateApplicationForm('Phone', value || '')}
                        />
                        <TextField
                            label="Your Experience"
                            required
                            placeholder="Example: 4 Years"
                            value={applicationForm.Experience}
                            onChange={(_, value) => updateApplicationForm('Experience', value || '')}
                        />
                        <TextField
                            label="Expected Salary"
                            required
                            placeholder="Example: 8-10 LPA"
                            value={applicationForm.ExpectedSalary}
                            onChange={(_, value) => updateApplicationForm('ExpectedSalary', value || '')}
                        />
                        <TextField
                            label="Cover Letter"
                            multiline
                            rows={8}
                            placeholder="Enter your cover letter..."
                            value={applicationForm.CoverLetter}
                            onChange={(_, value) => updateApplicationForm('CoverLetter', value || '')}
                        />

                        <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 20 } }}>
                            <PrimaryButton
                                text={isApplying ? 'Submitting...' : 'Apply'}
                                disabled={isApplying}
                                onClick={submitApplication}
                            />
                            <DefaultButton
                                text="Cancel"
                                disabled={isApplying}
                                onClick={() => setIsApplyPanelOpen(false)}
                            />
                        </Stack>
                    </Stack>
                    </>
                )}
            </Panel>
        </div>
    );
};

const headerStyle: React.CSSProperties = {
    textAlign: 'left',
    padding: 12,
    backgroundColor: '#f3f2f1',
    borderBottom: '1px solid #ddd',
    fontWeight: 600,
    whiteSpace: 'nowrap'
};

const cellStyle: React.CSSProperties = {
    padding: 12,
    verticalAlign: 'top',
    borderBottom: '1px solid #edebe9'
};

const jobInfoStyle: React.CSSProperties = {
    background: '#f3f2f1',
    padding: 16,
    borderRadius: 4
};

export default JobApplicationPortalComponent;