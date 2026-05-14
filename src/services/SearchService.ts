import { SPFI } from "@pnp/sp";
import "@pnp/sp/search";

export interface ISearchResult {
    title: string;
    path: string;
    fileType: string;
    author: string;
    modified: string;
    created: string;
    siteTitle: string;
    contentClass: string;
}

export default class SearchService {

    private sp: SPFI;

    constructor(sp: SPFI) {
        this.sp = sp;
    }

    public async searchTenantContent(
        searchText: string,
    ): Promise<ISearchResult[]> {

        const sanitizedSearch = searchText?.trim();

        // Build a query that ONLY searches the Title property
        let query = sanitizedSearch ? `Title:${sanitizedSearch}*` : "*";

        const results = await this.sp.search({
            Querytext: query,
            RowLimit: 30,
            SelectProperties: [
                "Title",
                "Path",
                "FileType",
                "Author",
                "LastModifiedTime",
                "Created",
                "SiteTitle",
                "ContentClass",
                "OriginalPath"
            ],
            SortList: [
                {
                    Property: "Created",
                    Direction: 1 // 1 = Descending (Newest first)
                }
            ]
        });

        return results.PrimarySearchResults.map((item: any) => ({
            title: item.Title || "Untitled",
            path: item.Path || item.OriginalPath || "",
            fileType: item.FileType || (item.contentclass === "STS_Web" ? "Site" : "Item"),
            author: item.Author || "System",
            modified: item.LastModifiedTime ? item.LastModifiedTime.toString() : "N/A",
            created: item.Created ? item.Created.toString() : "N/A",
            siteTitle: item.SiteTitle || "Tenant Content",
            contentClass: item.contentclass || ""
        }));
    }
}