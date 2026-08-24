import * as React from 'react';
import { Web } from 'sp-pnp-js';
import {
    Stack,
    Text,
    Icon,
    Link,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    SearchBox,
    DefaultButton,
    FontWeights,
    mergeStyleSets,
    useTheme
} from '@fluentui/react';

import {
    ColumnDef,
    flexRender,
    getCoreRowModel,
    getFilteredRowModel,
    getPaginationRowModel,
    getSortedRowModel,
    SortingState,
    useReactTable
} from '@tanstack/react-table';

interface IDocument {
    Id: number;
    Title: string;
    FileRef: string;
    FileDirRef: string;
    FileLeafRef: string;
    userName: string;
}

const DocumentPortalComponent = (allProps: any) => {

    const theme = useTheme();

    const [documents, setDocuments] = React.useState<IDocument[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);
    const [error, setError] = React.useState<string>('');
    const [sorting, setSorting] = React.useState<SortingState>([]);
    const [globalFilter, setGlobalFilter] = React.useState<string>('');

    const classNames = React.useMemo(() => mergeStyleSets({

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

        headerStack: {
            marginBottom: '20px'
        },

        title: {
            fontSize: '24px',
            fontWeight: FontWeights.semibold,
            color: theme.palette.neutralPrimary
        },

        subtitle: {
            fontSize: '14px',
            color: theme.palette.neutralSecondary,
            marginTop: '4px'
        },

        toolbar: {
            width: '100%',
            marginBottom: '16px'
        },

        searchBox: {
            width: '280px'
        },

        countBadge: {
            backgroundColor: theme.palette.themeLighterAlt,
            color: theme.palette.themePrimary,
            padding: '6px 12px',
            borderRadius: '16px',
            fontSize: '13px',
            fontWeight: FontWeights.semibold
        },

        detailsListWrapper: {
            width: '100%',
            maxWidth: '100%',
            overflowX: 'auto',
            border: `1px solid ${theme.palette.neutralLight}`,
            borderRadius: theme.effects.roundedCorner4
        },

        table: {
            width: '100%',
            borderCollapse: 'collapse',
            tableLayout: 'fixed'
        },

        tableHeader: {
            backgroundColor: theme.palette.neutralLighterAlt,
            borderBottom: `1px solid ${theme.palette.neutralLight}`,
            padding: '12px 16px',
            textAlign: 'left',
            fontSize: '13px',
            fontWeight: FontWeights.semibold,
            color: theme.palette.neutralPrimary,
            cursor: 'pointer',
            userSelect: 'none'
        },

        tableCell: {
            padding: '13px 16px',
            borderBottom: `1px solid ${theme.palette.neutralLighter}`,
            fontSize: '14px',
            color: theme.palette.neutralPrimary,
            verticalAlign: 'middle'
        },

        tableRow: {
            selectors: {
                '&:hover': {
                    backgroundColor: theme.palette.neutralLighterAlt
                }
            }
        },

        documentCell: {
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            minWidth: 0
        },

        documentIcon: {
            fontSize: '18px',
            color: theme.palette.themePrimary,
            flexShrink: 0
        },

        documentLink: {
            color: theme.palette.themePrimary,
            fontWeight: FontWeights.semibold,
            textDecoration: 'none',
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            selectors: {
                '&:hover': {
                    textDecoration: 'underline'
                }
            }
        },

        userCell: {
            display: 'flex',
            alignItems: 'center',
            gap: '8px'
        },

        userAvatar: {
            width: '28px',
            height: '28px',
            borderRadius: '50%',
            backgroundColor: theme.palette.themePrimary,
            color: theme.palette.white,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '12px',
            fontWeight: FontWeights.semibold,
            flexShrink: 0
        },

        urlLink: {
            display: 'inline-flex',
            alignItems: 'center',
            gap: '6px',
            color: theme.palette.themePrimary,
            textDecoration: 'none',
            cursor: 'pointer',
            selectors: {
                '&:hover': {
                    textDecoration: 'underline'
                }
            }
        },

        sortIcon: {
            marginLeft: '6px',
            fontSize: '11px'
        },

        pagination: {
            marginTop: '16px'
        },

        pageInfo: {
            fontSize: '13px',
            color: theme.palette.neutralSecondary
        },

        emptyState: {
            padding: '50px 20px',
            textAlign: 'center',
            color: theme.palette.neutralSecondary
        },

        emptyIcon: {
            fontSize: '36px',
            color: theme.palette.neutralTertiary,
            marginBottom: '12px'
        },

        loadingContainer: {
            minHeight: '250px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
        }

    }), [theme]);

    // --------------------------------------------------
    // Load Documents
    // --------------------------------------------------

    const loadDocuments = async () => {

        try {

            setLoading(true);
            setError('');

            const baseUrl = allProps.allProps.SiteUrl;
            const documentListId = allProps.allProps.DocumentList;
            const userRegistrationDetailsListId =
                allProps.allProps.UserRegistrationDetailsList;

            const web = new Web(baseUrl);

            // Get registered users
            const userRegistrationDetails = await web.lists
                .getById(userRegistrationDetailsListId)
                .items
                .select(
                    'Id',
                    'Title',
                    'ResumeId'
                )
                .getAll();

            // Get only files - folders are excluded using FSObjType eq 0
            const fetchedDocuments = await web.lists
                .getById(documentListId)
                .items
                .select(
                    'Id',
                    'Title',
                    'FileRef',
                    'FileDirRef',
                    'FileLeafRef',
                    'FSObjType'
                )
                .filter('FSObjType eq 0')
                .getAll();

            // Create lookup map
            const userMap = new Map<number, string>(
                userRegistrationDetails.map((user: any) => [
                    user.ResumeId,
                    user.Title
                ])
            );

            // Combine document and user data
            const updatedDocuments: IDocument[] =
                fetchedDocuments.map((doc: any) => ({
                    Id: doc.Id,
                    Title: doc.Title,
                    FileRef: doc.FileRef,
                    FileDirRef: doc.FileDirRef,
                    FileLeafRef: doc.FileLeafRef,
                    userName: userMap.get(doc.Id) || 'Unknown User'
                }));

            setDocuments(updatedDocuments);

        } catch (err) {

            console.error('Error fetching documents:', err);

            setError(
                'Unable to load documents. Please try again.'
            );

        } finally {

            setLoading(false);

        }
    };

    React.useEffect(() => {
        loadDocuments();
    }, []);

    // --------------------------------------------------
    // TanStack Columns
    // --------------------------------------------------

    const columns = React.useMemo<ColumnDef<IDocument>[]>(() => [

        {
            accessorKey: 'FileLeafRef',
            header: 'Document Name',

            cell: ({ row }) => {

                const document = row.original;

                const fileName = document.FileLeafRef || 'Unnamed Document';

                const extension =
                    fileName.split('.').pop()?.toLowerCase();

                let iconName = 'Document';

                if (extension === 'pdf') {
                    iconName = 'PDF';
                } else if (
                    extension === 'doc' ||
                    extension === 'docx'
                ) {
                    iconName = 'WordDocument';
                } else if (
                    extension === 'xls' ||
                    extension === 'xlsx'
                ) {
                    iconName = 'ExcelDocument';
                } else if (
                    extension === 'ppt' ||
                    extension === 'pptx'
                ) {
                    iconName = 'PowerPointDocument';
                }

                return (
                    <div className={classNames.documentCell}>

                        <Icon
                            iconName={iconName}
                            className={classNames.documentIcon}
                        />

                        <Link
                            href={document.FileRef}
                            target="_blank"
                            data-interception="off"
                            rel="noopener noreferrer"
                            className={classNames.documentLink}
                            title={fileName}
                        >
                            {fileName}
                        </Link>

                    </div>
                );
            }
        },

        {
            accessorKey: 'userName',
            header: 'User',

            cell: ({ row }) => {

                const userName =
                    row.original.userName || 'Unknown User';

                return (
                    <div className={classNames.userCell}>

                        <div className={classNames.userAvatar}>
                            {userName
                                .charAt(0)
                                .toUpperCase()}
                        </div>

                        <span>
                            {userName}
                        </span>

                    </div>
                );
            }
        },

        {
            accessorKey: 'FileRef',
            header: 'URL',

            cell: ({ row }) => {

                const url = row.original.FileRef;

                return (
                    <Link
                        href={url}
                        target="_blank"
                        data-interception="off"
                        rel="noopener noreferrer"
                        className={classNames.urlLink}
                    >
                        <Icon iconName="Link" />
                        <span>Open Document</span>
                    </Link>
                );
            }
        }

    ], [classNames]);

    // --------------------------------------------------
    // TanStack Table
    // --------------------------------------------------

    const table = useReactTable({
        data: documents,
        columns,

        state: {
            sorting,
            globalFilter
        },

        onSortingChange: setSorting,
        onGlobalFilterChange: setGlobalFilter,

        getCoreRowModel: getCoreRowModel(),
        getSortedRowModel: getSortedRowModel(),
        getFilteredRowModel: getFilteredRowModel(),
        getPaginationRowModel: getPaginationRowModel(),

        initialState: {
            pagination: {
                pageSize: 10
            }
        }
    });

    // --------------------------------------------------
    // Loading
    // --------------------------------------------------

    if (loading) {

        return (
            <div className={classNames.container}>

                <div className={classNames.loadingContainer}>

                    <Spinner
                        size={SpinnerSize.large}
                        label="Loading documents..."
                    />

                </div>

            </div>
        );
    }

    // --------------------------------------------------
    // Error
    // --------------------------------------------------

    if (error) {

        return (
            <div className={classNames.container}>

                <MessageBar
                    messageBarType={MessageBarType.error}
                >
                    {error}
                </MessageBar>

            </div>
        );
    }

    // --------------------------------------------------
    // Main UI
    // --------------------------------------------------

    return (
        <div className={classNames.container}>

            {/* Header */}

            <Stack
                horizontal
                horizontalAlign="space-between"
                verticalAlign="center"
                className={classNames.headerStack}
            >

                <Stack>

                    <Text className={classNames.title}>
                        Document Portal
                    </Text>

                    <Text className={classNames.subtitle}>
                        View and access uploaded documents
                    </Text>

                </Stack>

                <Text className={classNames.countBadge}>
                    {documents.length} Document
                    {documents.length !== 1 ? 's' : ''}
                </Text>

            </Stack>


            {/* Toolbar */}

            <Stack
                horizontal
                horizontalAlign="space-between"
                verticalAlign="center"
                className={classNames.toolbar}
            >

                <SearchBox
                    placeholder="Search documents..."
                    value={globalFilter}
                    onChange={(_, newValue) => {
                        setGlobalFilter(newValue || '');
                    }}
                    className={classNames.searchBox}
                />

                <DefaultButton
                    text="Refresh"
                    iconProps={{
                        iconName: 'Refresh'
                    }}
                    onClick={loadDocuments}
                />

            </Stack>


            {/* Table */}

            {table.getRowModel().rows.length === 0 ? (

                <div className={classNames.emptyState}>

                    <Icon
                        iconName="Document"
                        className={classNames.emptyIcon}
                    />

                    <Text variant="large">
                        No documents found
                    </Text>

                    <Text>
                        Try changing your search criteria.
                    </Text>

                </div>

            ) : (

                <div className={classNames.detailsListWrapper}>

                    <table className={classNames.table}>

                        <thead>

                            {table.getHeaderGroups().map(
                                headerGroup => (

                                    <tr key={headerGroup.id}>

                                        {headerGroup.headers.map(
                                            header => (

                                                <th
                                                    key={header.id}
                                                    className={
                                                        classNames.tableHeader
                                                    }
                                                    onClick={
                                                        header.column.getToggleSortingHandler()
                                                    }
                                                >

                                                    {flexRender(
                                                        header.column
                                                            .columnDef.header,
                                                        header.getContext()
                                                    )}

                                                    {{
                                                        asc: ' ▲',
                                                        desc: ' ▼'
                                                    }[
                                                        header.column.getIsSorted() as string
                                                    ] || ''}

                                                </th>

                                            )
                                        )}

                                    </tr>
                                )
                            )}

                        </thead>

                        <tbody>

                            {table
                                .getRowModel()
                                .rows
                                .map(row => (

                                    <tr
                                        key={row.original.Id}
                                        className={classNames.tableRow}
                                    >

                                        {row
                                            .getVisibleCells()
                                            .map(cell => (

                                                <td
                                                    key={cell.id}
                                                    className={
                                                        classNames.tableCell
                                                    }
                                                >

                                                    {flexRender(
                                                        cell.column
                                                            .columnDef.cell,
                                                        cell.getContext()
                                                    )}

                                                </td>

                                            ))}

                                    </tr>

                                ))}

                        </tbody>

                    </table>

                </div>

            )}


            {/* Pagination */}

            {table.getPageCount() > 0 && (

                <Stack
                    horizontal
                    horizontalAlign="space-between"
                    verticalAlign="center"
                    className={classNames.pagination}
                >

                    <Text className={classNames.pageInfo}>

                        Page{' '}
                        {table.getState().pagination.pageIndex + 1}
                        {' '}of{' '}
                        {table.getPageCount()}

                    </Text>


                    <Stack
                        horizontal
                        tokens={{ childrenGap: 8 }}
                    >

                        <DefaultButton
                            text="Previous"
                            disabled={
                                !table.getCanPreviousPage()
                            }
                            onClick={() =>
                                table.previousPage()
                            }
                        />

                        <DefaultButton
                            text="Next"
                            disabled={
                                !table.getCanNextPage()
                            }
                            onClick={() =>
                                table.nextPage()
                            }
                        />

                    </Stack>

                </Stack>

            )}

        </div>
    );
};

export default DocumentPortalComponent;