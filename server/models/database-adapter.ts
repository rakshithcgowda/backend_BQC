// Database adapter that switches between SQLite (local) and Postgres (production)
import { Database } from './database.js';

// Use SQLite for local development when PostgreSQL connection fails
// PostgreSQL will only be loaded if explicitly enabled to avoid connection errors
let selectedDatabase: any;

// Check if we should use PostgreSQL (only if POSTGRES_URL is set and USE_POSTGRES is true)
// For local development without a running PostgreSQL instance, use SQLite
if (process.env.POSTGRES_URL && process.env.USE_POSTGRES === 'true') {
  // Dynamically import PostgreSQL only when needed (lazy load)
  // This prevents the module from being loaded and initializing the connection
  import('./postgres-database.js').then((module) => {
    selectedDatabase = module.getPostgresDatabase();
    console.log('📊 Using PostgreSQL database');
  }).catch((error) => {
    console.warn('⚠️  Failed to load PostgreSQL, falling back to SQLite:', error);
    selectedDatabase = new Database();
  });
  // For now, use SQLite until PostgreSQL loads
  selectedDatabase = new Database();
} else {
  // Default to SQLite for local development
  selectedDatabase = new Database();
  if (process.env.POSTGRES_URL) {
    console.log('✅ Using SQLite database (set USE_POSTGRES=true in .env.local to use PostgreSQL)');
  } else {
    console.log('✅ Using SQLite database for local development');
  }
}

export const database = selectedDatabase;
