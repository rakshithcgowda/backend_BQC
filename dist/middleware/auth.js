"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.authenticateToken = authenticateToken;
exports.authenticateTokenVercel = authenticateTokenVercel;
exports.generateToken = generateToken;
const jsonwebtoken_1 = __importDefault(require("jsonwebtoken"));
const JWT_SECRET = process.env.JWT_SECRET || 'your-secret-key-change-in-production';
function authenticateToken(req, res, next) {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];
    if (!token) {
        return res.status(401).json({
            success: false,
            message: 'Access token required'
        });
    }
    jsonwebtoken_1.default.verify(token, JWT_SECRET, (err, decoded) => {
        if (err) {
            return res.status(403).json({
                success: false,
                message: 'Invalid or expired token'
            });
        }
        req.userId = decoded.userId;
        next();
    });
}
// Vercel-compatible auth function
async function authenticateTokenVercel(req) {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];
    if (!token) {
        return {
            success: false,
            message: 'Access token required'
        };
    }
    try {
        const decoded = jsonwebtoken_1.default.verify(token, JWT_SECRET);
        return {
            success: true,
            userId: decoded.userId
        };
    }
    catch (err) {
        return {
            success: false,
            message: 'Invalid or expired token'
        };
    }
}
function generateToken(userId) {
    return jsonwebtoken_1.default.sign({ userId }, JWT_SECRET, { expiresIn: '24h' });
}
