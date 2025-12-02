import express from 'express';
import { database } from '../models/database-adapter.js';
import { authenticateToken } from '../middleware/auth.js';
const router = express.Router();
// All admin routes require authentication
router.use(authenticateToken);
// Check if user is admin (you can implement proper admin role checking)
const isAdmin = (req, res, next) => {
    // For now, we'll allow all authenticated users to access admin features
    // In production, you should implement proper role-based access control
    next();
};
router.use(isAdmin);
// Get BQC statistics overview
router.get('/stats-overview', async (req, res) => {
    try {
        const { startDate, endDate, groupName } = req.query;
        const stats = await database.getBQCStats({
            startDate: startDate,
            endDate: endDate,
            groupName: groupName
        });
        res.json({
            success: true,
            data: stats
        });
    }
    catch (error) {
        console.error('Admin stats error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch statistics'
        });
    }
});
// Get BQC statistics by group
router.get('/stats-groups', async (req, res) => {
    try {
        const { startDate, endDate, groupName } = req.query;
        const groupStats = await database.getBQCGroupStats({
            startDate: startDate,
            endDate: endDate,
            groupName: groupName
        });
        res.json({
            success: true,
            data: groupStats
        });
    }
    catch (error) {
        console.error('Admin group stats error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch group statistics'
        });
    }
});
// Get BQC statistics by date range
router.get('/stats-date-range', async (req, res) => {
    try {
        const { startDate, endDate, groupBy = 'day' } = req.query;
        const dateStats = await database.getBQCDateRangeStats({
            startDate: startDate,
            endDate: endDate,
            groupBy: groupBy
        });
        res.json({
            success: true,
            data: dateStats
        });
    }
    catch (error) {
        console.error('Admin date range stats error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch date range statistics'
        });
    }
});
// Get all BQC entries with pagination and filters
router.get('/bqc-entries', async (req, res) => {
    try {
        const { page = 1, limit = 10, startDate, endDate, groupName, tenderType, search } = req.query;
        const entries = await database.getBQCEntries({
            page: parseInt(page),
            limit: parseInt(limit),
            startDate: startDate,
            endDate: endDate,
            groupName: groupName,
            tenderType: tenderType,
            search: search
        });
        res.json({
            success: true,
            data: entries
        });
    }
    catch (error) {
        console.error('Admin BQC entries error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch BQC entries'
        });
    }
});
// Get user statistics
router.get('/stats-users', async (req, res) => {
    try {
        const userStats = await database.getUserStats();
        res.json({
            success: true,
            data: userStats
        });
    }
    catch (error) {
        console.error('Admin user stats error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch user statistics'
        });
    }
});
// Get tender type statistics
router.get('/stats-tender-types', async (req, res) => {
    try {
        const { startDate, endDate, groupName } = req.query;
        const tenderTypeStats = await database.getTenderTypeStats({
            startDate: startDate,
            endDate: endDate,
            groupName: groupName
        });
        res.json({
            success: true,
            data: tenderTypeStats
        });
    }
    catch (error) {
        console.error('Admin tender type stats error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch tender type statistics'
        });
    }
});
// Get financial statistics
router.get('/stats-financial', async (req, res) => {
    try {
        const { startDate, endDate, groupName } = req.query;
        const financialStats = await database.getFinancialStats({
            startDate: startDate,
            endDate: endDate,
            groupName: groupName
        });
        res.json({
            success: true,
            data: financialStats
        });
    }
    catch (error) {
        console.error('Admin financial stats error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch financial statistics'
        });
    }
});
// Export BQC data
router.get('/export', async (req, res) => {
    try {
        const { startDate, endDate, groupName, format = 'csv' } = req.query;
        const exportData = await database.exportBQCData({
            startDate: startDate,
            endDate: endDate,
            groupName: groupName,
            format: format
        });
        if (format === 'csv') {
            res.setHeader('Content-Type', 'text/csv');
            res.setHeader('Content-Disposition', 'attachment; filename="bqc-data.csv"');
            res.send(exportData);
        }
        else {
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename="bqc-data.xlsx"');
            res.send(exportData);
        }
    }
    catch (error) {
        console.error('Admin export error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to export data'
        });
    }
});
// User management routes
router.get('/users', async (req, res) => {
    try {
        const users = await database.getAllUsers();
        res.json({
            success: true,
            data: users
        });
    }
    catch (error) {
        console.error('Admin users error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to fetch users'
        });
    }
});
router.post('/users', async (req, res) => {
    try {
        const { userId, action } = req.body;
        if (!userId || !action) {
            return res.status(400).json({
                success: false,
                message: 'User ID and action are required'
            });
        }
        if (action === 'approve') {
            if (!req.userId) {
                return res.status(401).json({
                    success: false,
                    message: 'Authentication required'
                });
            }
            await database.approveUser(userId, req.userId);
            res.json({
                success: true,
                message: 'User approved successfully'
            });
        }
        else if (action === 'reject') {
            await database.rejectUser(userId);
            res.json({
                success: true,
                message: 'User rejected and removed'
            });
        }
        else {
            return res.status(400).json({
                success: false,
                message: 'Invalid action. Use "approve" or "reject"'
            });
        }
    }
    catch (error) {
        console.error('Admin user management error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to manage user'
        });
    }
});
export default router;
//# sourceMappingURL=admin.js.map