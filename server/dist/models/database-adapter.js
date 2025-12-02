// Database adapter that switches between SQLite (local) and Postgres (production)
import { postgresDatabase } from './postgres-database.js';
// For now, we'll use Postgres for both local and production
// You can add SQLite back for local development if needed
export const database = postgresDatabase;
//# sourceMappingURL=database-adapter.js.map