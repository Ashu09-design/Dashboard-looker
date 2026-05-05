const jwt = require('jsonwebtoken');

const JWT_SECRET = process.env.JWT_SECRET || 'dashboard-looker-secret-key-change-me';
const TOKEN_EXPIRY = '30m';

/**
 * Middleware: verify JWT Bearer token on protected routes.
 * Attaches decoded payload to req.user on success.
 */
function verifyToken(req, res, next) {
    const authHeader = req.headers['authorization'];
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({ error: 'Access denied. No token provided.' });
    }

    const token = authHeader.split(' ')[1];
    try {
        const decoded = jwt.verify(token, JWT_SECRET);
        req.user = decoded;
        next();
    } catch (err) {
        if (err.name === 'TokenExpiredError') {
            return res.status(401).json({ error: 'Session expired. Please log in again.' });
        }
        return res.status(401).json({ error: 'Invalid token.' });
    }
}

/**
 * Generate a signed JWT for a given user payload.
 */
function signToken(payload) {
    return jwt.sign(payload, JWT_SECRET, { expiresIn: TOKEN_EXPIRY });
}

module.exports = { verifyToken, signToken, JWT_SECRET };
