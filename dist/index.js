"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
require("dotenv/config");
const dotenv_1 = require("dotenv");
// Load .env.local only in development
if (process.env.NODE_ENV !== 'production') {
    (0, dotenv_1.config)({ path: '.env.local' });
}
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const helmet_1 = __importDefault(require("helmet"));
const express_rate_limit_1 = __importDefault(require("express-rate-limit"));
const path_1 = __importDefault(require("path"));
const url_1 = require("url");
const auth_js_1 = __importDefault(require("./routes/auth.js"));
const bqc_js_1 = __importDefault(require("./routes/bqc.js"));
const admin_js_1 = __importDefault(require("./routes/admin.js"));
const __filename = (0, url_1.fileURLToPath)(import.meta.url);
const __dirname = path_1.default.dirname(__filename);
const app = (0, express_1.default)();
const PORT = process.env.PORT || 3002;
// Security middleware
app.use((0, helmet_1.default)());
// Rate limiting
const limiter = (0, express_rate_limit_1.default)({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 100, // limit each IP to 100 requests per windowMs
    message: {
        success: false,
        message: 'Too many requests from this IP, please try again later.'
    }
});
app.use('/api/', limiter);
// CORS configuration
app.use((0, cors_1.default)({
    origin: process.env.NODE_ENV === 'production'
        ? true // Allow all origins in production for Vercel
        : ['http://localhost:3000', 'http://127.0.0.1:3000'],
    credentials: true
}));
// Body parsing middleware
app.use(express_1.default.json({ limit: '10mb' }));
app.use(express_1.default.urlencoded({ extended: true, limit: '10mb' }));
// Health check endpoint
app.get('/health', (req, res) => {
    res.json({
        success: true,
        message: 'BQC Generator API is running',
        timestamp: new Date().toISOString()
    });
});
// API routes
app.use('/api/auth', auth_js_1.default);
app.use('/api/bqc', bqc_js_1.default);
app.use('/api/admin', admin_js_1.default);
// Serve static files from frontend dist in production
if (process.env.NODE_ENV === 'production') {
    const distPath = path_1.default.join(__dirname, '../../frontend/dist');
    app.use(express_1.default.static(distPath));
    // Serve index.html for all non-API routes (SPA routing)
    app.get(/^\/(?!api).*/, (req, res) => {
        res.sendFile(path_1.default.join(distPath, 'index.html'));
    });
}
// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Error:', err);
    res.status(err.status || 500).json({
        success: false,
        message: process.env.NODE_ENV === 'production'
            ? 'Internal server error'
            : err.message || 'Something went wrong'
    });
});
// 404 handler for API routes only (SPA handles frontend 404s)
app.use(/^\/api\/.*/, (req, res) => {
    res.status(404).json({
        success: false,
        message: 'API endpoint not found'
    });
});
// Start server
app.listen(PORT, () => {
    console.log(`🚀 BQC Generator API server running on port ${PORT}`);
    console.log(`📊 Health check: http://localhost:${PORT}/health`);
    console.log(`🔐 Environment: ${process.env.NODE_ENV || 'development'}`);
});
// Prevent process from exiting on unhandled errors; log them instead
process.on('unhandledRejection', (reason) => {
    console.error('Unhandled Rejection:', reason);
});
process.on('uncaughtException', (error) => {
    console.error('Uncaught Exception:', error);
});
// Graceful shutdown
process.on('SIGTERM', () => {
    console.log('SIGTERM received, shutting down gracefully');
    process.exit(0);
});
process.on('SIGINT', () => {
    console.log('SIGINT received, shutting down gracefully');
    process.exit(0);
});
exports.default = app;
