import 'dotenv/config';
import { config } from 'dotenv';
// Load .env.local only in development
if (process.env.NODE_ENV !== 'production') {
    config({ path: '.env.local' });
}
import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import path from 'path';
import { fileURLToPath } from 'url';
import authRoutes from './routes/auth.js';
import bqcRoutes from './routes/bqc.js';
import adminRoutes from './routes/admin.js';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = express();
const PORT = process.env.PORT || 3002;
// Security middleware
app.use(helmet());
// Rate limiting
const limiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 100, // limit each IP to 100 requests per windowMs
    message: {
        success: false,
        message: 'Too many requests from this IP, please try again later.'
    }
});
app.use('/api/', limiter);
// CORS configuration
app.use(cors({
    origin: process.env.NODE_ENV === 'production'
        ? true // Allow all origins in production for Vercel
        : ['http://localhost:3000', 'http://127.0.0.1:3000'],
    credentials: true
}));
// Body parsing middleware
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
// Health check endpoint
app.get('/health', (req, res) => {
    res.json({
        success: true,
        message: 'BQC Generator API is running',
        timestamp: new Date().toISOString()
    });
});
// API routes
app.use('/api/auth', authRoutes);
app.use('/api/bqc', bqcRoutes);
app.use('/api/admin', adminRoutes);
// Serve static files from frontend dist in production
if (process.env.NODE_ENV === 'production') {
    const distPath = path.join(__dirname, '../../frontend/dist');
    app.use(express.static(distPath));
    // Serve index.html for all non-API routes (SPA routing)
    app.get(/^\/(?!api).*/, (req, res) => {
        res.sendFile(path.join(distPath, 'index.html'));
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
export default app;
//# sourceMappingURL=index.js.map