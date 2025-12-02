import sqlite3 from 'sqlite3';
import path from 'path';
import fs from 'fs';

// Get application data directory for portable database storage
function getAppDataDir(): string {
  let appData: string;
  
  if (process.platform === 'win32') {
    appData = process.env.LOCALAPPDATA || path.join(process.env.USERPROFILE || '', 'AppData', 'Local');
  } else if (process.platform === 'darwin') {
    appData = path.join(process.env.HOME || '', 'Library', 'Application Support');
  } else {
    appData = process.env.XDG_DATA_HOME || path.join(process.env.HOME || '', '.local', 'share');
  }
  
  const appDir = path.join(appData, 'BQCGenerator');
  if (!fs.existsSync(appDir)) {
    fs.mkdirSync(appDir, { recursive: true });
  }
  
  return appDir;
}

const APP_DATA_DIR = getAppDataDir();
const DB_PATH = path.join(APP_DATA_DIR, 'user_data.db');

export class Database {
  private db: sqlite3.Database;

  constructor() {
    this.db = new sqlite3.Database(DB_PATH);
    this.setupDatabase();
  }

  private setupDatabase(): void {
    // Create users table
    this.db.run(`
      CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        email TEXT,
        full_name TEXT,
        is_approved INTEGER DEFAULT 1,
        approved_at TIMESTAMP,
        approved_by INTEGER,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      )
    `);

    // Create bqc_data table
    this.db.run(`
      CREATE TABLE IF NOT EXISTS bqc_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        ref_number TEXT NOT NULL,
        group_name TEXT,
        subject TEXT,
        tender_description TEXT,
        pr_reference TEXT,
        tender_type TEXT,
        evaluation_methodology TEXT,
        cec_estimate_incl_gst REAL,
        cec_date DATE,
        cec_estimate_excl_gst REAL,
        lots TEXT,
        quantity_supplied REAL,
        budget_details TEXT,
        tender_platform TEXT,
        scope_of_work TEXT,
        contract_period_months TEXT,
        contract_duration_years REAL,
        delivery_period TEXT,
        bid_validity_period TEXT,
        warranty_period TEXT,
        amc_period TEXT,
        payment_terms TEXT,
        manufacturer_types TEXT,
        supplying_capacity INTEGER,
        mse_relaxation INTEGER,
        similar_work_definition TEXT,
        annualized_value REAL,
        escalation_clause TEXT,
        divisibility TEXT,
        performance_security INTEGER,
        has_performance_security INTEGER,
        proposed_by TEXT,
        proposed_by_designation TEXT,
        recommended_by TEXT,
        recommended_by_designation TEXT,
        concurred_by TEXT,
        concurred_by_designation TEXT,
        approved_by TEXT,
        approved_by_designation TEXT,
        amc_value REAL,
        has_amc INTEGER,
        correction_factor REAL,
        o_m_value REAL,
        o_m_period TEXT,
        has_om INTEGER,
        additional_details TEXT,
        note_to TEXT,
        commercial_evaluation_method TEXT,
        has_experience_explanatory_note INTEGER,
        experience_explanatory_note TEXT,
        has_additional_explanatory_note INTEGER,
        additional_explanatory_note TEXT,
        has_financial_explanatory_note INTEGER,
        financial_explanatory_note TEXT,
        has_emd_explanatory_note INTEGER,
        emd_explanatory_note TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users (id)
      )
    `);

    // Add migration for contract_duration_years column if it doesn't exist
    this.db.run(`
      ALTER TABLE bqc_data ADD COLUMN contract_duration_years REAL DEFAULT 1
    `, (err) => {
      // Ignore error if column already exists
      if (err && !err.message.includes('duplicate column name')) {
        console.error('Error adding contract_duration_years column:', err);
      }
    });

    // Add migration for bid_validity_period column if it doesn't exist
    this.db.run(`
      ALTER TABLE bqc_data ADD COLUMN bid_validity_period TEXT
    `, (err) => {
      // Ignore error if column already exists
      if (err && !err.message.includes('duplicate column name')) {
        console.error('Error adding bid_validity_period column:', err);
      }
    });

    // Add migration for explanatory note fields
    const explanatoryNoteFields = [
      'has_experience_explanatory_note INTEGER',
      'experience_explanatory_note TEXT',
      'has_additional_explanatory_note INTEGER',
      'additional_explanatory_note TEXT',
      'has_financial_explanatory_note INTEGER',
      'financial_explanatory_note TEXT',
      'has_emd_explanatory_note INTEGER',
      'emd_explanatory_note TEXT',
      'has_past_performance_explanatory_note INTEGER',
      'past_performance_explanatory_note TEXT',
      'past_performance_mse_relaxation INTEGER'
    ];

    explanatoryNoteFields.forEach(field => {
      this.db.run(`
        ALTER TABLE bqc_data ADD COLUMN ${field}
      `, (err) => {
        // Ignore error if column already exists
        if (err && !err.message.includes('duplicate column name')) {
          console.error(`Error adding ${field} column:`, err);
        }
      });
    });

    // Add migration for user approval columns
    const userApprovalColumns = [
      'is_approved INTEGER DEFAULT 1',
      'approved_at TIMESTAMP',
      'approved_by INTEGER'
    ];

    userApprovalColumns.forEach(field => {
      this.db.run(`
        ALTER TABLE users ADD COLUMN ${field}
      `, (err) => {
        // Ignore error if column already exists
        if (err && !err.message.includes('duplicate column name')) {
          console.error(`Error adding ${field} column:`, err);
        }
      });
    });

    // Add migration for missing columns that might not exist
    const missingColumns = [
      'evaluation_methodology TEXT',
      'subject TEXT',
      'tender_description TEXT',
      'pr_reference TEXT',
      'tender_type TEXT',
      'cec_estimate_incl_gst REAL',
      'cec_date DATE',
      'cec_estimate_excl_gst REAL',
      'lots TEXT',
      'quantity_supplied REAL',
      'budget_details TEXT',
      'tender_platform TEXT',
      'scope_of_work TEXT',
      'contract_period_months TEXT',
      'delivery_period TEXT',
      'bid_validity_period TEXT',
      'warranty_period TEXT',
      'amc_period TEXT',
      'payment_terms TEXT',
      'manufacturer_types TEXT',
      'supplying_capacity INTEGER',
      'mse_relaxation INTEGER',
      'similar_work_definition TEXT',
      'annualized_value REAL',
      'escalation_clause TEXT',
      'divisibility TEXT',
      'performance_security INTEGER',
      'has_performance_security INTEGER',
      'proposed_by TEXT',
      'proposed_by_designation TEXT',
      'recommended_by TEXT',
      'recommended_by_designation TEXT',
      'concurred_by TEXT',
      'concurred_by_designation TEXT',
      'approved_by TEXT',
      'approved_by_designation TEXT',
      'amc_value REAL',
      'has_amc INTEGER',
      'correction_factor REAL',
      'o_m_value REAL',
      'o_m_period TEXT',
      'has_om INTEGER',
      'additional_details TEXT',
      'note_to TEXT',
      'commercial_evaluation_method TEXT'
    ];

    missingColumns.forEach(field => {
      this.db.run(`
        ALTER TABLE bqc_data ADD COLUMN ${field}
      `, (err) => {
        // Ignore error if column already exists
        if (err && !err.message.includes('duplicate column name')) {
          console.error(`Error adding ${field} column:`, err);
        }
      });
    });
  }

  // User operations
  async createUser(userData: {
    username: string;
    password: string;
    email: string;
    fullName: string;
  }): Promise<number> {
    return new Promise((resolve, reject) => {
      const { username, password, email, fullName } = userData;
      this.db.run(
        'INSERT INTO users (username, password, email, full_name) VALUES (?, ?, ?, ?)',
        [username, password, email, fullName],
        function(err) {
          if (err) {
            reject(err);
          } else {
            resolve(this.lastID);
          }
        }
      );
    });
  }

  async getUserByUsername(username: string): Promise<any> {
    return new Promise((resolve, reject) => {
      this.db.get(
        'SELECT * FROM users WHERE username = ?',
        [username],
        (err, row) => {
          if (err) {
            reject(err);
          } else {
            resolve(row);
          }
        }
      );
    });
  }

  async getUserById(id: number): Promise<any> {
    return new Promise((resolve, reject) => {
      this.db.get(
        'SELECT id, username, email, full_name, is_approved, approved_at, approved_by, created_at FROM users WHERE id = ?',
        [id],
        (err, row) => {
          if (err) {
            reject(err);
          } else {
            resolve(row);
          }
        }
      );
    });
  }

  // User approval methods
  async getPendingUsers(): Promise<any[]> {
    return new Promise((resolve, reject) => {
      this.db.all(
        'SELECT id, username, email, full_name, created_at FROM users WHERE is_approved = 0 OR is_approved IS NULL ORDER BY created_at ASC',
        (err, rows) => {
          if (err) {
            reject(err);
          } else {
            resolve(rows || []);
          }
        }
      );
    });
  }

  async approveUser(userId: number, approvedBy: number): Promise<void> {
    return new Promise((resolve, reject) => {
      this.db.run(
        'UPDATE users SET is_approved = 1, approved_at = CURRENT_TIMESTAMP, approved_by = ? WHERE id = ?',
        [approvedBy, userId],
        (err) => {
          if (err) {
            reject(err);
          } else {
            resolve();
          }
        }
      );
    });
  }

  async rejectUser(userId: number): Promise<void> {
    return new Promise((resolve, reject) => {
      this.db.run(
        'DELETE FROM users WHERE id = ? AND (is_approved = 0 OR is_approved IS NULL)',
        [userId],
        (err) => {
          if (err) {
            reject(err);
          } else {
            resolve();
          }
        }
      );
    });
  }

  async getAllUsers(): Promise<any[]> {
    return new Promise((resolve, reject) => {
      this.db.all(
        `SELECT 
          u.id, 
          u.username, 
          u.email, 
          u.full_name, 
          u.is_approved, 
          u.approved_at, 
          u.created_at,
          approver.username as approved_by_username
        FROM users u
        LEFT JOIN users approver ON u.approved_by = approver.id
        ORDER BY u.created_at DESC`,
        (err, rows) => {
          if (err) {
            reject(err);
          } else {
            resolve(rows || []);
          }
        }
      );
    });
  }

  // BQC data operations
  async saveBQCData(userId: number, bqcData: any): Promise<number> {
    return new Promise((resolve, reject) => {
      // Check if record exists
      this.db.get(
        'SELECT id FROM bqc_data WHERE user_id = ? AND ref_number = ?',
        [userId, bqcData.refNumber],
        (err, existingRecord: any) => {
          if (err) {
            reject(err);
            return;
          }

          const manufacturerTypesJson = JSON.stringify(bqcData.manufacturerTypes || []);
          const commercialEvaluationMethodJson = JSON.stringify(bqcData.commercialEvaluationMethod || []);
          
          if (existingRecord) {
            // Update existing record
            this.db.run(`
              UPDATE bqc_data SET
                group_name = ?, subject = ?, tender_description = ?, pr_reference = ?,
                tender_type = ?, evaluation_methodology = ?, cec_estimate_incl_gst = ?, cec_date = ?,
                cec_estimate_excl_gst = ?, lots = ?, quantity_supplied = ?, budget_details = ?, tender_platform = ?,
                scope_of_work = ?, contract_period_months = ?, contract_duration_years = ?, delivery_period = ?, bid_validity_period = ?,
                warranty_period = ?, amc_period = ?, payment_terms = ?,
                manufacturer_types = ?, supplying_capacity = ?, mse_relaxation = ?,
                similar_work_definition = ?, annualized_value = ?, escalation_clause = ?,
                divisibility = ?, performance_security = ?, has_performance_security = ?, proposed_by = ?, proposed_by_designation = ?,
                recommended_by = ?, recommended_by_designation = ?, concurred_by = ?, concurred_by_designation = ?, approved_by = ?, approved_by_designation = ?,
                amc_value = ?, has_amc = ?, correction_factor = ?,
                o_m_value = ?, o_m_period = ?, has_om = ?, additional_details = ?, note_to = ?, commercial_evaluation_method = ?,
                has_experience_explanatory_note = ?, experience_explanatory_note = ?, has_additional_explanatory_note = ?, additional_explanatory_note = ?,
                has_financial_explanatory_note = ?, financial_explanatory_note = ?, has_emd_explanatory_note = ?, emd_explanatory_note = ?,
                has_past_performance_explanatory_note = ?, past_performance_explanatory_note = ?, past_performance_mse_relaxation = ?, updated_at = CURRENT_TIMESTAMP
              WHERE id = ?
            `, [
              bqcData.groupName, bqcData.subject, bqcData.tenderDescription, bqcData.prReference,
              bqcData.tenderType, bqcData.evaluationMethodology, bqcData.cecEstimateInclGst, bqcData.cecDate,
              bqcData.cecEstimateExclGst, JSON.stringify(bqcData.lots || []), bqcData.quantitySupplied, bqcData.budgetDetails, bqcData.tenderPlatform,
              bqcData.scopeOfWork, bqcData.contractPeriodMonths, bqcData.contractDurationYears, bqcData.deliveryPeriod, bqcData.bidValidityPeriod,
              bqcData.warrantyPeriod, bqcData.amcPeriod, bqcData.paymentTerms,
              manufacturerTypesJson, bqcData.supplyingCapacity?.final || 0, bqcData.mseRelaxation ? 1 : 0,
              bqcData.similarWorkDefinition, bqcData.annualizedValue, bqcData.escalationClause,
              bqcData.divisibility, bqcData.performanceSecurity, bqcData.hasPerformanceSecurity ? 1 : 0, bqcData.proposedBy, bqcData.proposedByDesignation,
              bqcData.recommendedBy, bqcData.recommendedByDesignation, bqcData.concurredBy, bqcData.concurredByDesignation, bqcData.approvedBy, bqcData.approvedByDesignation,
              bqcData.amcValue, bqcData.hasAmc ? 1 : 0, bqcData.correctionFactor,
              bqcData.omValue || 0, bqcData.omPeriod || '', bqcData.hasOm ? 1 : 0,
              bqcData.additionalDetails, bqcData.noteTo, commercialEvaluationMethodJson,
              bqcData.hasExperienceExplanatoryNote ? 1 : 0, bqcData.experienceExplanatoryNote || '',
              bqcData.hasAdditionalExplanatoryNote ? 1 : 0, bqcData.additionalExplanatoryNote || '',
              bqcData.hasFinancialExplanatoryNote ? 1 : 0, bqcData.financialExplanatoryNote || '',
              bqcData.hasEMDExplanatoryNote ? 1 : 0, bqcData.emdExplanatoryNote || '',
              bqcData.hasPastPerformanceExplanatoryNote ? 1 : 0, bqcData.pastPerformanceExplanatoryNote || '',
              bqcData.pastPerformanceMseRelaxation ? 1 : 0, existingRecord.id
            ], function(err) {
              if (err) {
                reject(err);
              } else {
                resolve(existingRecord.id);
              }
            });
          } else {
            // Insert new record
            this.db.run(`
              INSERT INTO bqc_data (
                user_id, ref_number, group_name, subject, tender_description, pr_reference,
                tender_type, evaluation_methodology, cec_estimate_incl_gst, cec_date, cec_estimate_excl_gst,
                lots, quantity_supplied, budget_details, tender_platform,                 scope_of_work, contract_period_months, contract_duration_years,
                delivery_period, bid_validity_period, warranty_period, amc_period, payment_terms,
                manufacturer_types, supplying_capacity, mse_relaxation,
                similar_work_definition, annualized_value, escalation_clause,
                divisibility, performance_security, has_performance_security, proposed_by, proposed_by_designation, recommended_by, recommended_by_designation,
                concurred_by, concurred_by_designation, approved_by, approved_by_designation, amc_value, has_amc, correction_factor,
                o_m_value, o_m_period, has_om, additional_details, note_to, commercial_evaluation_method,
                has_experience_explanatory_note, experience_explanatory_note, has_additional_explanatory_note, additional_explanatory_note,
                has_financial_explanatory_note, financial_explanatory_note, has_emd_explanatory_note, emd_explanatory_note,
                has_past_performance_explanatory_note, past_performance_explanatory_note, past_performance_mse_relaxation
              ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            `, [
              userId, bqcData.refNumber, bqcData.groupName, bqcData.subject, bqcData.tenderDescription,
              bqcData.prReference, bqcData.tenderType, bqcData.evaluationMethodology, bqcData.cecEstimateInclGst,
              bqcData.cecDate, bqcData.cecEstimateExclGst, JSON.stringify(bqcData.lots || []), bqcData.quantitySupplied, bqcData.budgetDetails,
              bqcData.tenderPlatform, bqcData.scopeOfWork, bqcData.contractPeriodMonths, bqcData.contractDurationYears,
              bqcData.deliveryPeriod, bqcData.bidValidityPeriod, bqcData.warrantyPeriod, bqcData.amcPeriod,
              bqcData.paymentTerms, manufacturerTypesJson, bqcData.supplyingCapacity?.final || 0,
              bqcData.mseRelaxation ? 1 : 0, bqcData.similarWorkDefinition,
              bqcData.annualizedValue, bqcData.escalationClause, bqcData.divisibility,
              bqcData.performanceSecurity, bqcData.hasPerformanceSecurity ? 1 : 0, bqcData.proposedBy, bqcData.proposedByDesignation, bqcData.recommendedBy, bqcData.recommendedByDesignation,
              bqcData.concurredBy, bqcData.concurredByDesignation, bqcData.approvedBy, bqcData.approvedByDesignation, bqcData.amcValue,
              bqcData.hasAmc ? 1 : 0, bqcData.correctionFactor,               bqcData.omValue || 0,
              bqcData.omPeriod || '', bqcData.hasOm ? 1 : 0, bqcData.additionalDetails, bqcData.noteTo, commercialEvaluationMethodJson,
              bqcData.hasExperienceExplanatoryNote ? 1 : 0, bqcData.experienceExplanatoryNote || '',
              bqcData.hasAdditionalExplanatoryNote ? 1 : 0, bqcData.additionalExplanatoryNote || '',
              bqcData.hasFinancialExplanatoryNote ? 1 : 0, bqcData.financialExplanatoryNote || '',
              bqcData.hasEMDExplanatoryNote ? 1 : 0, bqcData.emdExplanatoryNote || '',
              bqcData.hasPastPerformanceExplanatoryNote ? 1 : 0, bqcData.pastPerformanceExplanatoryNote || '',
              bqcData.pastPerformanceMseRelaxation ? 1 : 0
            ], function(err) {
              if (err) {
                reject(err);
              } else {
                resolve(this.lastID);
              }
            });
          }
        }
      );
    });
  }

  async getBQCData(userId: number, id: number): Promise<any> {
    return new Promise((resolve, reject) => {
      this.db.get(
        'SELECT * FROM bqc_data WHERE id = ? AND user_id = ?',
        [id, userId],
        (err, row: any) => {
          if (err) {
            reject(err);
          } else {
            if (row) {
              // Parse manufacturer types
              try {
                row.manufacturer_types = JSON.parse(row.manufacturer_types || '[]');
              } catch {
                row.manufacturer_types = [];
              }
              
              // Parse commercial evaluation method
              try {
                row.commercial_evaluation_method = JSON.parse(row.commercial_evaluation_method || '[]');
              } catch {
                row.commercial_evaluation_method = [];
              }
              
              // Parse lots data
              try {
                row.lots = JSON.parse(row.lots || '[]');
              } catch {
                row.lots = [];
              }
            }
            resolve(row);
          }
        }
      );
    });
  }

  // Alias method for compatibility with load API
  async getBQCDataById(id: number, userId: number): Promise<any> {
    return this.getBQCData(userId, id);
  }

  async listBQCData(userId: number): Promise<any[]> {
    return new Promise((resolve, reject) => {
      this.db.all(
        'SELECT id, ref_number, tender_description, created_at FROM bqc_data WHERE user_id = ? ORDER BY created_at DESC',
        [userId],
        (err, rows) => {
          if (err) {
            reject(err);
          } else {
            resolve(rows || []);
          }
        }
      );
    });
  }

  async deleteBQCData(userId: number, id: number): Promise<void> {
    return new Promise((resolve, reject) => {
      this.db.run(
        'DELETE FROM bqc_data WHERE id = ? AND user_id = ?',
        [id, userId],
        (err) => {
          if (err) {
            reject(err);
          } else {
            resolve();
          }
        }
      );
    });
  }

  // Admin statistics methods
  async getBQCStats(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any> {
    return new Promise((resolve, reject) => {
      let query = `
        SELECT 
          COUNT(*) as totalBQCs,
          COUNT(DISTINCT user_id) as totalUsers,
          SUM(cec_estimate_incl_gst) as totalValue,
          AVG(cec_estimate_incl_gst) as avgValue,
          COUNT(CASE WHEN tender_type = 'Goods' THEN 1 END) as goodsCount,
          COUNT(CASE WHEN tender_type = 'Service' THEN 1 END) as serviceCount,
          COUNT(CASE WHEN tender_type = 'Works' THEN 1 END) as worksCount,
          COUNT(CASE WHEN evaluation_methodology = 'least cash outflow' THEN 1 END) as leastCashOutflowCount,
          COUNT(CASE WHEN evaluation_methodology = 'Lot-wise' THEN 1 END) as lotWiseCount
        FROM bqc_data
        WHERE 1=1
      `;
      
      const params: any[] = [];
      
      if (filters.startDate) {
        query += ' AND DATE(created_at) >= ?';
        params.push(filters.startDate);
      }
      
      if (filters.endDate) {
        query += ' AND DATE(created_at) <= ?';
        params.push(filters.endDate);
      }
      
      if (filters.groupName) {
        query += ' AND group_name = ?';
        params.push(filters.groupName);
      }

      this.db.get(query, params, (err, row) => {
        if (err) {
          reject(err);
        } else {
          resolve(row);
        }
      });
    });
  }

  async getBQCGroupStats(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any[]> {
    return new Promise((resolve, reject) => {
      let query = `
        SELECT 
          group_name,
          COUNT(*) as count,
          SUM(cec_estimate_incl_gst) as totalValue,
          AVG(cec_estimate_incl_gst) as avgValue,
          COUNT(CASE WHEN tender_type = 'Goods' THEN 1 END) as goodsCount,
          COUNT(CASE WHEN tender_type = 'Service' THEN 1 END) as serviceCount,
          COUNT(CASE WHEN tender_type = 'Works' THEN 1 END) as worksCount
        FROM bqc_data
        WHERE 1=1
      `;
      
      const params: any[] = [];
      
      if (filters.startDate) {
        query += ' AND DATE(created_at) >= ?';
        params.push(filters.startDate);
      }
      
      if (filters.endDate) {
        query += ' AND DATE(created_at) <= ?';
        params.push(filters.endDate);
      }
      
      if (filters.groupName) {
        query += ' AND group_name = ?';
        params.push(filters.groupName);
      }
      
      query += ' GROUP BY group_name ORDER BY count DESC';

      this.db.all(query, params, (err, rows) => {
        if (err) {
          reject(err);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  async getBQCDateRangeStats(filters: {
    startDate?: string;
    endDate?: string;
    groupBy: 'day' | 'week' | 'month';
  }): Promise<any[]> {
    return new Promise((resolve, reject) => {
      let dateFormat: string;
      switch (filters.groupBy) {
        case 'day':
          dateFormat = '%Y-%m-%d';
          break;
        case 'week':
          dateFormat = '%Y-%W';
          break;
        case 'month':
          dateFormat = '%Y-%m';
          break;
        default:
          dateFormat = '%Y-%m-%d';
      }

      let query = `
        SELECT 
          strftime('${dateFormat}', created_at) as period,
          COUNT(*) as count,
          SUM(cec_estimate_incl_gst) as totalValue
        FROM bqc_data
        WHERE 1=1
      `;
      
      const params: any[] = [];
      
      if (filters.startDate) {
        query += ' AND DATE(created_at) >= ?';
        params.push(filters.startDate);
      }
      
      if (filters.endDate) {
        query += ' AND DATE(created_at) <= ?';
        params.push(filters.endDate);
      }
      
      query += ` GROUP BY strftime('${dateFormat}', created_at) ORDER BY period`;

      this.db.all(query, params, (err, rows) => {
        if (err) {
          reject(err);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  async getBQCEntries(filters: {
    page: number;
    limit: number;
    startDate?: string;
    endDate?: string;
    groupName?: string;
    tenderType?: string;
    search?: string;
  }): Promise<{ entries: any[]; total: number; totalPages: number }> {
    return new Promise((resolve, reject) => {
      let whereClause = 'WHERE 1=1';
      const params: any[] = [];
      
      if (filters.startDate) {
        whereClause += ' AND DATE(b.created_at) >= ?';
        params.push(filters.startDate);
      }
      
      if (filters.endDate) {
        whereClause += ' AND DATE(b.created_at) <= ?';
        params.push(filters.endDate);
      }
      
      if (filters.groupName) {
        whereClause += ' AND b.group_name = ?';
        params.push(filters.groupName);
      }
      
      if (filters.tenderType) {
        whereClause += ' AND b.tender_type = ?';
        params.push(filters.tenderType);
      }
      
      if (filters.search) {
        whereClause += ' AND (b.ref_number LIKE ? OR b.tender_description LIKE ? OR b.subject LIKE ?)';
        const searchTerm = `%${filters.search}%`;
        params.push(searchTerm, searchTerm, searchTerm);
      }

      // Get total count
      const countQuery = `
        SELECT COUNT(*) as total
        FROM bqc_data b
        ${whereClause}
      `;

      this.db.get(countQuery, params, (err, countRow: any) => {
        if (err) {
          reject(err);
          return;
        }

        const total = countRow.total;
        const totalPages = Math.ceil(total / filters.limit);
        const offset = (filters.page - 1) * filters.limit;

        // Get paginated entries
        const entriesQuery = `
          SELECT 
            b.id,
            b.ref_number,
            b.group_name,
            b.subject,
            b.tender_description,
            b.tender_type,
            b.cec_estimate_incl_gst,
            b.created_at,
            u.username,
            u.full_name
          FROM bqc_data b
          LEFT JOIN users u ON b.user_id = u.id
          ${whereClause}
          ORDER BY b.created_at DESC
          LIMIT ? OFFSET ?
        `;

        this.db.all(entriesQuery, [...params, filters.limit, offset], (err, rows) => {
          if (err) {
            reject(err);
          } else {
            resolve({
              entries: rows || [],
              total,
              totalPages
            });
          }
        });
      });
    });
  }

  async getUserStats(): Promise<any> {
    return new Promise((resolve, reject) => {
      const query = `
        SELECT 
          COUNT(*) as totalUsers,
          COUNT(CASE WHEN DATE(created_at) >= DATE('now', '-30 days') THEN 1 END) as newUsersLast30Days,
          COUNT(CASE WHEN DATE(created_at) >= DATE('now', '-7 days') THEN 1 END) as newUsersLast7Days
        FROM users
      `;

      this.db.get(query, (err, row) => {
        if (err) {
          reject(err);
        } else {
          resolve(row);
        }
      });
    });
  }

  async getTenderTypeStats(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any[]> {
    return new Promise((resolve, reject) => {
      let query = `
        SELECT 
          tender_type,
          COUNT(*) as count,
          SUM(cec_estimate_incl_gst) as totalValue,
          AVG(cec_estimate_incl_gst) as avgValue
        FROM bqc_data
        WHERE 1=1
      `;
      
      const params: any[] = [];
      
      if (filters.startDate) {
        query += ' AND DATE(created_at) >= ?';
        params.push(filters.startDate);
      }
      
      if (filters.endDate) {
        query += ' AND DATE(created_at) <= ?';
        params.push(filters.endDate);
      }
      
      if (filters.groupName) {
        query += ' AND group_name = ?';
        params.push(filters.groupName);
      }
      
      query += ' GROUP BY tender_type ORDER BY count DESC';

      this.db.all(query, params, (err, rows) => {
        if (err) {
          reject(err);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  async getFinancialStats(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any> {
    return new Promise((resolve, reject) => {
      let query = `
        SELECT 
          SUM(cec_estimate_incl_gst) as totalValue,
          AVG(cec_estimate_incl_gst) as avgValue,
          MIN(cec_estimate_incl_gst) as minValue,
          MAX(cec_estimate_incl_gst) as maxValue,
          COUNT(CASE WHEN cec_estimate_incl_gst < 100000 THEN 1 END) as under1Lakh,
          COUNT(CASE WHEN cec_estimate_incl_gst >= 100000 AND cec_estimate_incl_gst < 1000000 THEN 1 END) as between1LakhAnd10Lakh,
          COUNT(CASE WHEN cec_estimate_incl_gst >= 1000000 AND cec_estimate_incl_gst < 10000000 THEN 1 END) as between10LakhAnd1Crore,
          COUNT(CASE WHEN cec_estimate_incl_gst >= 10000000 THEN 1 END) as above1Crore
        FROM bqc_data
        WHERE 1=1
      `;
      
      const params: any[] = [];
      
      if (filters.startDate) {
        query += ' AND DATE(created_at) >= ?';
        params.push(filters.startDate);
      }
      
      if (filters.endDate) {
        query += ' AND DATE(created_at) <= ?';
        params.push(filters.endDate);
      }
      
      if (filters.groupName) {
        query += ' AND group_name = ?';
        params.push(filters.groupName);
      }

      this.db.get(query, params, (err, row) => {
        if (err) {
          reject(err);
        } else {
          resolve(row);
        }
      });
    });
  }

  async exportBQCData(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
    format: 'csv' | 'excel';
  }): Promise<string> {
    return new Promise((resolve, reject) => {
      let query = `
        SELECT 
          b.id,
          b.ref_number,
          b.group_name,
          b.subject,
          b.tender_description,
          b.tender_type,
          b.evaluation_methodology,
          b.cec_estimate_incl_gst,
          b.cec_estimate_excl_gst,
          b.lots,
          b.annualized_value,
          b.similar_work_definition,
          b.created_at,
          u.username,
          u.full_name
        FROM bqc_data b
        LEFT JOIN users u ON b.user_id = u.id
        WHERE 1=1
      `;
      
      const params: any[] = [];
      
      if (filters.startDate) {
        query += ' AND DATE(b.created_at) >= ?';
        params.push(filters.startDate);
      }
      
      if (filters.endDate) {
        query += ' AND DATE(b.created_at) <= ?';
        params.push(filters.endDate);
      }
      
      if (filters.groupName) {
        query += ' AND b.group_name = ?';
        params.push(filters.groupName);
      }
      
      query += ' ORDER BY b.created_at DESC';

      this.db.all(query, params, (err, rows) => {
        if (err) {
          reject(err);
        } else {
          if (filters.format === 'csv') {
            // Convert to CSV with lot-wise data
            const csvHeader = 'ID,Ref Number,Group,Subject,Tender Description,Tender Type,Evaluation Methodology,CEC (Incl GST),CEC (Excl GST),Annualized Value,Similar Work Definition,Lot Data,Created At,Username,Full Name\n';
            const csvRows = (rows || []).map((row: any) => {
              // Parse lot data to extract individual lot information
              let lotDataString = '';
              try {
                const lots = row.lots ? JSON.parse(row.lots) : [];
                if (lots && lots.length > 0) {
                  const lotDetails = lots.map((lot: any, index: number) => {
                    return `Lot ${index + 1}: CEC=${lot.cecEstimateInclGst || 0}, EMD=${lot.emdAmount || 0}, Similar Work A=${lot.similarWorksOptionA || 0}, Similar Work B=${lot.similarWorksOptionB || 0}, Similar Work C=${lot.similarWorksOptionC || 0}, Contract Period=${lot.contractPeriodText || lot.contractPeriodMonths || 0} months`;
                  }).join('; ');
                  lotDataString = lotDetails;
                } else {
                  lotDataString = 'No lots defined';
                }
              } catch (error) {
                lotDataString = 'Error parsing lot data';
              }
              
              return `${row.id},"${row.ref_number}","${row.group_name}","${row.subject}","${row.tender_description}","${row.tender_type}","${row.evaluation_methodology}",${row.cec_estimate_incl_gst},${row.cec_estimate_excl_gst},${row.annualized_value || 0},"${row.similar_work_definition || ''}","${lotDataString}","${row.created_at}","${row.username}","${row.full_name}"`;
            }).join('\n');
            resolve(csvHeader + csvRows);
          } else {
            // For Excel, we'll return CSV for now (you can implement proper Excel export later)
            const csvHeader = 'ID,Ref Number,Group,Subject,Tender Description,Tender Type,Evaluation Methodology,CEC (Incl GST),CEC (Excl GST),Annualized Value,Similar Work Definition,Lot Data,Created At,Username,Full Name\n';
            const csvRows = (rows || []).map((row: any) => {
              // Parse lot data to extract individual lot information
              let lotDataString = '';
              try {
                const lots = row.lots ? JSON.parse(row.lots) : [];
                if (lots && lots.length > 0) {
                  const lotDetails = lots.map((lot: any, index: number) => {
                    return `Lot ${index + 1}: CEC=${lot.cecEstimateInclGst || 0}, EMD=${lot.emdAmount || 0}, Similar Work A=${lot.similarWorksOptionA || 0}, Similar Work B=${lot.similarWorksOptionB || 0}, Similar Work C=${lot.similarWorksOptionC || 0}, Contract Period=${lot.contractPeriodText || lot.contractPeriodMonths || 0} months`;
                  }).join('; ');
                  lotDataString = lotDetails;
                } else {
                  lotDataString = 'No lots defined';
                }
              } catch (error) {
                lotDataString = 'Error parsing lot data';
              }
              
              return `${row.id},"${row.ref_number}","${row.group_name}","${row.subject}","${row.tender_description}","${row.tender_type}","${row.evaluation_methodology}",${row.cec_estimate_incl_gst},${row.cec_estimate_excl_gst},${row.annualized_value || 0},"${row.similar_work_definition || ''}","${lotDataString}","${row.created_at}","${row.username}","${row.full_name}"`;
            }).join('\n');
            resolve(csvHeader + csvRows);
          }
        }
      });
    });
  }

  // Admin statistics methods - alias for compatibility
  async getAdminStatsOverview(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any> {
    return this.getBQCStats(filters);
  }

  async getAdminStatsGroups(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any[]> {
    return this.getBQCGroupStats(filters);
  }

  async getAdminStatsDateRange(filters: {
    startDate?: string;
    endDate?: string;
    groupBy: 'day' | 'week' | 'month';
  }): Promise<any[]> {
    return this.getBQCDateRangeStats(filters);
  }

  async getAdminStatsUsers(): Promise<any> {
    return this.getUserStats();
  }

  async getAdminStatsTenderTypes(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any[]> {
    return this.getTenderTypeStats(filters);
  }

  async getAdminStatsFinancial(filters: {
    startDate?: string;
    endDate?: string;
    groupName?: string;
  } = {}): Promise<any> {
    return this.getFinancialStats(filters);
  }

  async getAdminBQCEntries(filters: {
    page: number;
    limit: number;
    startDate?: string;
    endDate?: string;
    groupName?: string;
    tenderType?: string;
    search?: string;
  }): Promise<{ entries: any[]; total: number; totalPages: number }> {
    return this.getBQCEntries(filters);
  }

  async getAdminExportData(filters: {
    format: string;
    startDate?: string;
    endDate?: string;
    groupName?: string;
  }): Promise<string> {
    return this.exportBQCData({
      ...filters,
      format: filters.format as 'csv' | 'excel'
    });
  }

  close(): void {
    this.db.close();
  }
}

export const database = new Database();
