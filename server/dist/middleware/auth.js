import jwt from 'jsonwebtoken';
const JWT_SECRET = process.env.JWT_SECRET || 'your-secret-key-change-in-production';
export function authenticateToken(req, res, next) {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];
    if (!token) {
        return res.status(401).json({
            success: false,
            message: 'Access token required'
        });
    }
    jwt.verify(token, JWT_SECRET, (err, decoded) => {
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
export async function authenticateTokenVercel(req) {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];
    if (!token) {
        return {
            success: false,
            message: 'Access token required'
        };
    }
    try {
        const decoded = jwt.verify(token, JWT_SECRET);
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
export function generateToken(userId) {
    return jwt.sign({ userId }, JWT_SECRET, { expiresIn: '24h' });
}
//# sourceMappingURL=auth.js.map