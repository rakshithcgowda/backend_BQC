declare class PostgresDatabase {
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
    getPendingUsers(): Promise<any[]>;
    approveUser(userId: number, approvedBy: number): Promise<void>;
    rejectUser(userId: number): Promise<void>;
    getAllUsers(): Promise<any[]>;
    saveBQCData(userId: number, bqcData: any): Promise<number>;
    getBQCData(userId: number, id: number): Promise<any>;
    getBQCDataById(id: number, userId: number): Promise<any>;
    listBQCData(userId: number): Promise<any[]>;
    deleteBQCData(userId: number, id: number): Promise<void>;
    getAdminStatsOverview(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any>;
    getAdminStatsGroups(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any[]>;
    getAdminStatsDateRange(filters: {
        startDate?: string;
        endDate?: string;
        groupBy: 'day' | 'week' | 'month';
    }): Promise<any[]>;
    getAdminStatsUsers(): Promise<any>;
    getAdminStatsTenderTypes(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any[]>;
    getAdminStatsFinancial(filters?: {
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<any>;
    getAdminBQCEntries(filters: {
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
    getAdminExportData(filters: {
        format: string;
        startDate?: string;
        endDate?: string;
        groupName?: string;
    }): Promise<string>;
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
}
export declare const postgresDatabase: PostgresDatabase;
export {};
//# sourceMappingURL=postgres-database.d.ts.map