import express from 'express';
import bcrypt from 'bcryptjs';
import { database } from '../models/database-adapter.js';
import { generateToken, authenticateToken, AuthRequest } from '../middleware/auth.js';

const router = express.Router();

// Register
router.post('/register', async (req, res) => {
  try {
    const { username, password, email, fullName } = req.body;

    // Validate input
    if (!username || !password || !email || !fullName) {
      return res.status(400).json({
        success: false,
        message: 'All fields are required'
      });
    }

    if (password.length < 6) {
      return res.status(400).json({
        success: false,
        message: 'Password must be at least 6 characters long'
      });
    }

    // Check if user already exists
    const existingUser = await database.getUserByUsername(username);
    if (existingUser) {
      return res.status(400).json({
        success: false,
        message: 'Username already exists'
      });
    }

    // Hash password
    const hashedPassword = await bcrypt.hash(password, 10);

    // Create user
    const userId = await database.createUser({
      username,
      password: hashedPassword,
      email,
      fullName
    });

    const user = await database.getUserById(userId);

    res.status(201).json({
      success: true,
      data: {
        user: {
          id: user.id,
          username: user.username,
          email: user.email,
          fullName: user.full_name
        }
      },
      message: 'User registered successfully'
    });
  } catch (error) {
    console.error('Registration error:', error);
    res.status(500).json({
      success: false,
      message: 'Internal server error'
    });
  }
});

// Login
router.post('/login', async (req, res) => {
  try {
    console.log('Login attempt:', { username: req.body.username, timestamp: new Date().toISOString() });
    
    const { username, password } = req.body;

    // Validate input
    if (!username || !password) {
      console.log('Login validation failed: missing credentials');
      return res.status(400).json({
        success: false,
        message: 'Username and password are required'
      });
    }

    // Test database connection first
    try {
      await database.getUserByUsername('test');
    } catch (dbError) {
      console.error('Database connection test failed:', dbError);
      return res.status(500).json({
        success: false,
        message: 'Database connection failed',
        ...(process.env.NODE_ENV === 'development' && {
          error: dbError instanceof Error ? dbError.message : 'Database error'
        })
      });
    }

    // Find user
    const user = await database.getUserByUsername(username);
    if (!user) {
      console.log('Login failed: user not found:', username);
      return res.status(401).json({
        success: false,
        message: 'Invalid credentials'
      });
    }

    // Verify password
    const isValidPassword = await bcrypt.compare(password, user.password);
    if (!isValidPassword) {
      console.log('Login failed: invalid password for user:', username);
      return res.status(401).json({
        success: false,
        message: 'Invalid credentials'
      });
    }

    // Generate token
    const token = generateToken(user.id);
    console.log('Login successful for user:', username);

    res.json({
      success: true,
      data: {
        user: {
          id: user.id,
          username: user.username,
          email: user.email,
          fullName: user.full_name
        },
        token
      },
      message: 'Login successful'
    });
  } catch (error) {
    console.error('Login error:', error);
    console.error('Error details:', {
      message: error instanceof Error ? error.message : 'Unknown error',
      stack: error instanceof Error ? error.stack : undefined,
      name: error instanceof Error ? error.name : undefined,
      timestamp: new Date().toISOString()
    });
    res.status(500).json({
      success: false,
      message: 'Internal server error',
      ...(process.env.NODE_ENV === 'development' && {
        error: error instanceof Error ? error.message : 'Unknown error'
      })
    });
  }
});

// Get current user
router.get('/me', authenticateToken, async (req: AuthRequest, res) => {
  try {
    const user = await database.getUserById(req.userId!);
    if (!user) {
      return res.status(404).json({
        success: false,
        message: 'User not found'
      });
    }

    res.json({
      success: true,
      data: {
        id: user.id,
        username: user.username,
        email: user.email,
        fullName: user.full_name
      }
    });
  } catch (error) {
    console.error('Get user error:', error);
    res.status(500).json({
      success: false,
      message: 'Internal server error'
    });
  }
});

// Logout (client-side token removal, but we can track it server-side if needed)
router.post('/logout', authenticateToken, (req, res) => {
  res.json({
    success: true,
    message: 'Logout successful'
  });
});

export default router;
