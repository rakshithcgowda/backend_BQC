"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.database = void 0;
// Database adapter that switches between SQLite (local) and Postgres (production)
const postgres_database_js_1 = require("./postgres-database.js");
// For now, we'll use Postgres for both local and production
// You can add SQLite back for local development if needed
exports.database = postgres_database_js_1.postgresDatabase;
