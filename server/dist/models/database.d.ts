declare class Database {
    private db;
    constructor();
    private setupDatabase;
    createUser(userData: {
        username: string;
        password: string;
        email: string;
        fullName: string;
    }): Promise<number>;
    getUserByUsername(username: string): Promise<any>;
    getUserById(id: number): Promise<any>;
    saveBQCData(userId: number, bqcData: any): Promise<number>;
    getBQCData(userId: number, id: number): Promise<any>;
    getBQCDataById(id: number, userId: number): Promise<any>;
    listBQCData(userId: number): Promise<any[]>;
    deleteBQCData(userId: number, id: number): Promise<void>;
    getBQCStats(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any>;
    getBQCGroupStats(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any[]>;
    getBQCDateRangeStats(filters: {
        startDate?: string;
        endDate?: string;
        groupBy: 'day' | 'week' | 'month';
    }): Promise<any[]>;
    getBQCEntries(filters: {
        page: number;
        limit: number;
        startDate?: string;
        endDate?: string;
        groupName?: string;
        tenderType?: string;
        search?: string;
    }): Promise<{
        entries: any[];
        total: number;
        totalPages: number;
    }>;
    getUserStats(): Promise<any>;
    getTenderTypeStats(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any[]>;
    getFinancialStats(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any>;
    exportBQCData(filters: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
        format: 'csv' | 'excel';
    }): Promise<string>;
    close(): void;
}
export declare const database: Database;
export {};
//# sourceMappingURL=database.d.ts.map