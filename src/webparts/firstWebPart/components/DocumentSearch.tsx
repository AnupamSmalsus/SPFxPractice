import * as React from "react";
import { useState } from "react";
import styles from "./DocumentSearch.module.scss"

import { ISearchResult } from "../../../services/SearchService";

export interface ILibrarySearchProps {
    searchService: {
        searchTenantContent: (
            searchText: string
        ) => Promise<ISearchResult[]>;
    };
}

const DocumentSearch: React.FC<ILibrarySearchProps> = ({ searchService }) => {
    const [searchText, setSearchText] = useState<string>("");
    const [results, setResults] = useState<ISearchResult[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(false);

    const onSearch = async (): Promise<void> => {
        setIsLoading(true);

        try {
            const searchResults = await searchService.searchTenantContent(
                searchText.trim(),
            );

            setResults(searchResults);
        } catch (error) {
            console.error("Search error:", error);
            setResults([]);
        } finally {
            setIsLoading(false);
        }
    };

    const getFileIcon = (fileType: string): string => {
        switch (fileType?.toLowerCase()) {
            case "pdf": return "📕";
            case "docx": return "📘";
            case "xlsx": return "📗";
            case "pptx": return "📙";
            case "site": return "🌐";
            case "folder": return "📁";
            default: return "📄";
        }
    };

    return (
        <div className={styles.searchContainer}>
            <div className={styles.header}>
                <h2>OOTB Search</h2>
                <p>Explore all content and metadata across your tenant</p>
            </div>

            <div className={styles.searchBoxWrapper}>
                <input
                    type="text"
                    placeholder="Search anything (files, sites, items)..."
                    value={searchText}
                    onChange={(e) => setSearchText(e.target.value)}
                    onKeyDown={(e) => {
                        if (e.key === "Enter") {
                            void onSearch();
                        }
                    }}
                />

                <button onClick={() => { void onSearch(); }} disabled={isLoading}>
                    {isLoading ? "Exploring..." : "Explore"}
                </button>
            </div>



            {isLoading && (
                <div className={styles.loading}>
                    <h3>Exploring your tenant...</h3>
                    <p>Fetching universal metadata</p>
                </div>
            )}

            {!isLoading && results.length === 0 && (
                <div className={styles.emptyState}>
                    <h3>No Results Found</h3>
                    <p>Try broadening your search term or selecting 'ALL'.</p>
                </div>
            )}

            <div className={styles.resultsGrid}>
                {results.map((item, index) => (
                    <div className={styles.resultCard} key={index}>
                        <div className={styles.cardHeader}>
                            <div className={styles.fileIcon}>
                                {getFileIcon(item.fileType)}
                            </div>
                            <div className={styles.fileInfo}>
                                <h3>{item.title}</h3>
                            </div>
                        </div>

                        <div className={styles.cardBody}>
                            <div className={styles.metaItem}>
                                <strong>Category</strong>
                                <span>{item.fileType.toUpperCase()}</span>
                            </div>
                            <div className={styles.metaItem}>
                                <strong>Source Site</strong>
                                <span>{item.siteTitle}</span>
                            </div>
                            <div className={styles.metaItem}>
                                <strong>Author</strong>
                                <span>{item.author}</span>
                            </div>
                            <div className={styles.metaItem}>
                                <strong>Created</strong>
                                <span>
                                    {item.created !== "N/A"
                                        ? new Date(item.created).toLocaleDateString()
                                        : "N/A"}
                                </span>
                            </div>
                        </div>

                        <div className={styles.cardFooter}>
                            <a href={item.path} target="_blank" rel="noreferrer">
                                Open Result
                            </a>
                        </div>
                    </div>
                ))}
            </div>
        </div>
    );
};

export default DocumentSearch;