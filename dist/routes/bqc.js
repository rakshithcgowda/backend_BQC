"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const database_adapter_js_1 = require("../models/database-adapter.js");
const auth_js_1 = require("../middleware/auth.js");
const docx_1 = require("docx");
const htmlToWord_js_1 = require("../utils/htmlToWord.js");
const router = express_1.default.Router();
// All BQC routes require authentication
router.use(auth_js_1.authenticateToken);
// Save BQC data
router.post('/save', async (req, res) => {
    try {
        const bqcData = req.body;
        const userId = req.userId;
        // Only validate critical fields - other fields will be filled with defaults
        if (!bqcData.refNumber || bqcData.refNumber.trim() === '') {
            return res.status(400).json({
                success: false,
                message: 'Reference number is required'
            });
        }
        const id = await database_adapter_js_1.database.saveBQCData(userId, bqcData);
        res.json({
            success: true,
            data: { id },
            message: 'BQC data saved successfully'
        });
    }
    catch (error) {
        console.error('Save BQC error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to save BQC data'
        });
    }
});
// Load BQC data by ID
router.get('/load/:id', async (req, res) => {
    try {
        const id = parseInt(req.params.id);
        const userId = req.userId;
        if (isNaN(id)) {
            return res.status(400).json({
                success: false,
                message: 'Invalid ID'
            });
        }
        const bqcData = await database_adapter_js_1.database.getBQCData(userId, id);
        if (!bqcData) {
            return res.status(404).json({
                success: false,
                message: 'BQC data not found'
            });
        }
        // Convert database format to frontend format
        const formattedData = {
            id: bqcData.id,
            refNumber: bqcData.ref_number,
            groupName: bqcData.group_name,
            subject: bqcData.subject || '',
            tenderDescription: bqcData.tender_description,
            prReference: bqcData.pr_reference,
            tenderType: bqcData.tender_type,
            evaluationMethodology: bqcData.evaluation_methodology || 'least cash outflow',
            cecEstimateInclGst: bqcData.cec_estimate_incl_gst,
            cecDate: bqcData.cec_date,
            cecEstimateExclGst: bqcData.cec_estimate_excl_gst,
            lots: bqcData.lots || [],
            quantitySupplied: bqcData.quantity_supplied,
            budgetDetails: bqcData.budget_details,
            tenderPlatform: bqcData.tender_platform,
            scopeOfWork: bqcData.scope_of_work,
            contractPeriodMonths: bqcData.contract_period_months,
            contractDurationYears: bqcData.contract_duration_years || 1,
            deliveryPeriod: bqcData.delivery_period || '',
            warrantyPeriod: bqcData.warranty_period || '',
            amcPeriod: bqcData.amc_period || '',
            paymentTerms: bqcData.payment_terms || '',
            manufacturerTypes: bqcData.manufacturer_types || [],
            supplyingCapacity: {
                calculated: bqcData.supplying_capacity || 0,
                final: bqcData.supplying_capacity || 0,
                mseAdjusted: undefined
            },
            mseRelaxation: Boolean(bqcData.mse_relaxation),
            similarWorkDefinition: bqcData.similar_work_definition || '',
            annualizedValue: bqcData.annualized_value,
            escalationClause: bqcData.escalation_clause || '',
            divisibility: bqcData.divisibility,
            performanceSecurity: bqcData.performance_security,
            hasPerformanceSecurity: Boolean(bqcData.has_performance_security),
            proposedBy: bqcData.proposed_by,
            proposedByDesignation: bqcData.proposed_by_designation || '',
            recommendedBy: bqcData.recommended_by,
            recommendedByDesignation: bqcData.recommended_by_designation || '',
            concurredBy: bqcData.concurred_by,
            concurredByDesignation: bqcData.concurred_by_designation || '',
            approvedBy: bqcData.approved_by,
            approvedByDesignation: bqcData.approved_by_designation || '',
            amcValue: bqcData.amc_value,
            hasAmc: Boolean(bqcData.has_amc),
            correctionFactor: bqcData.correction_factor,
            omValue: bqcData.o_m_value || 0,
            omPeriod: bqcData.o_m_period || '',
            hasOm: Boolean(bqcData.has_om),
            additionalDetails: bqcData.additional_details || '',
            noteTo: bqcData.note_to || '',
            commercialEvaluationMethod: bqcData.commercial_evaluation_method || [],
            // Explanatory Notes
            hasExperienceExplanatoryNote: Boolean(bqcData.has_experience_explanatory_note),
            experienceExplanatoryNote: bqcData.experience_explanatory_note || '',
            hasAdditionalExplanatoryNote: Boolean(bqcData.has_additional_explanatory_note),
            additionalExplanatoryNote: bqcData.additional_explanatory_note || '',
            hasFinancialExplanatoryNote: Boolean(bqcData.has_financial_explanatory_note),
            financialExplanatoryNote: bqcData.financial_explanatory_note || '',
            hasEMDExplanatoryNote: Boolean(bqcData.has_emd_explanatory_note),
            emdExplanatoryNote: bqcData.emd_explanatory_note || '',
            hasPastPerformanceExplanatoryNote: Boolean(bqcData.has_past_performance_explanatory_note),
            pastPerformanceExplanatoryNote: bqcData.past_performance_explanatory_note || '',
            pastPerformanceMseRelaxation: Boolean(bqcData.past_performance_mse_relaxation)
        };
        res.json({
            success: true,
            data: formattedData
        });
    }
    catch (error) {
        console.error('Load BQC error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to load BQC data'
        });
    }
});
// List saved BQC entries
router.get('/list', async (req, res) => {
    try {
        const userId = req.userId;
        const entries = await database_adapter_js_1.database.listBQCData(userId);
        const formattedEntries = entries.map(entry => ({
            id: entry.id,
            refNumber: entry.ref_number,
            tenderDescription: entry.tender_description,
            createdAt: entry.created_at
        }));
        res.json({
            success: true,
            data: formattedEntries
        });
    }
    catch (error) {
        console.error('List BQC error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to load BQC entries'
        });
    }
});
// Delete BQC data
router.delete('/delete/:id', async (req, res) => {
    try {
        const id = parseInt(req.params.id);
        const userId = req.userId;
        if (isNaN(id)) {
            return res.status(400).json({
                success: false,
                message: 'Invalid ID'
            });
        }
        await database_adapter_js_1.database.deleteBQCData(userId, id);
        res.json({
            success: true,
            message: 'BQC data deleted successfully'
        });
    }
    catch (error) {
        console.error('Delete BQC error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to delete BQC data'
        });
    }
});
// Generate document
router.post('/generate', async (req, res) => {
    try {
        const { data: bqcData, format = 'docx' } = req.body;
        if (format !== 'docx') {
            return res.status(400).json({
                success: false,
                message: 'Only DOCX format is currently supported'
            });
        }
        // Calculate values for document
        const calculatePastPerformance = (quantitySupplied, mseRelaxation = false) => {
            const basePercentage = 0.30; // 30% of Quantity Supplied
            if (mseRelaxation) {
                return Math.round(quantitySupplied * basePercentage * (1 - 0.15)); // 15% relaxation
            }
            return Math.round(quantitySupplied * basePercentage);
        };
        const calculateEMD = (estimatedValue, tenderType) => {
            // Convert Cr to Lakhs for calculation (1 Cr = 100 Lakhs)
            const valueInLakhs = estimatedValue * 100;
            if (tenderType === 'Goods') {
                // Goods: 50L-100L = Nil, >100L = fixed amounts
                if (valueInLakhs >= 50 && valueInLakhs <= 100) {
                    return 0; // Nil
                }
                else if (valueInLakhs > 100 && valueInLakhs <= 500) {
                    return 2.5; // 2.5 Lakhs
                }
                else if (valueInLakhs > 500 && valueInLakhs <= 1000) {
                    return 5; // 5 Lakhs
                }
                else if (valueInLakhs > 1000 && valueInLakhs <= 1500) {
                    return 7.5; // 7.5 Lakhs
                }
                else if (valueInLakhs > 1500 && valueInLakhs <= 2500) {
                    return 10; // 10 Lakhs
                }
                else if (valueInLakhs > 2500) {
                    return 20; // 20 Lakhs
                }
                return 0; // For values < 50L
            }
            if (tenderType === 'Service') {
                // Service: 50L-100L = 1L (inclusive), >100L = fixed amounts
                if (valueInLakhs >= 50 && valueInLakhs <= 100) {
                    return 1; // 1 Lakh
                }
                else if (valueInLakhs > 100 && valueInLakhs <= 500) {
                    return 2.5; // 2.5 Lakhs
                }
                else if (valueInLakhs > 500 && valueInLakhs <= 1000) {
                    return 5; // 5 Lakhs
                }
                else if (valueInLakhs > 1000 && valueInLakhs <= 1500) {
                    return 7.5; // 7.5 Lakhs
                }
                else if (valueInLakhs > 1500 && valueInLakhs <= 2500) {
                    return 10; // 10 Lakhs
                }
                else if (valueInLakhs > 2500) {
                    return 20; // 20 Lakhs
                }
                return 0; // For values < 50L
            }
            if (tenderType === 'Works') {
                // Works: 50L-100L = 1L (inclusive), >100L = fixed amounts
                if (valueInLakhs >= 50 && valueInLakhs <= 100) {
                    return 1; // 1 Lakh
                }
                else if (valueInLakhs > 100 && valueInLakhs <= 500) {
                    return 2.5; // 2.5 Lakhs
                }
                else if (valueInLakhs > 500 && valueInLakhs <= 1000) {
                    return 5; // 5 Lakhs
                }
                else if (valueInLakhs > 1000 && valueInLakhs <= 1500) {
                    return 7.5; // 7.5 Lakhs
                }
                else if (valueInLakhs > 1500 && valueInLakhs <= 2500) {
                    return 10; // 10 Lakhs
                }
                else if (valueInLakhs > 2500) {
                    return 20; // 20 Lakhs
                }
                return 0; // For values < 50L
            }
            return 0; // Default case
        };
        const calculateTurnover = (data) => {
            let basePercentage = 0.3;
            if (data.divisibility === 'Divisible') {
                basePercentage = 0.3 * (1 + (data.correctionFactor || 0));
            }
            let baseAmount = 0;
            // For lot-wise, calculate total CEC including GST minus AMC
            if (data.evaluationMethodology === 'Lot-wise' && data.lots) {
                baseAmount = data.lots.reduce((sum, lot) => {
                    const lotCEC = lot.cecEstimateInclGst || 0;
                    const lotAMC = (lot.hasAmc && lot.amcValue && lot.amcValue > 0) ? lot.amcValue : 0;
                    return sum + (lotCEC - lotAMC);
                }, 0);
            }
            else {
                // For least cash outflow, use CEC including GST minus AMC
                const cecInclGst = data.cecEstimateInclGst || 0;
                const amcValue = (data.hasAmc && data.amcValue && data.amcValue > 0) ? data.amcValue : 0;
                baseAmount = cecInclGst - amcValue;
            }
            // Calculate turnover requirement
            const turnoverAmount = basePercentage * baseAmount;
            // Always apply annualization based on contract duration (divide by contract period)
            const contractDurationYears = data.contractDurationYears || 1;
            return turnoverAmount / contractDurationYears;
        };
        const calculateExperienceRequirements = (data) => {
            // Apply correction factor if divisible
            let optionAPercent = 0.4;
            let optionBPercent = 0.5;
            let optionCPercent = 0.8;
            if (data.divisibility === 'Divisible') {
                const correctionFactor = data.correctionFactor || 0;
                optionAPercent = 0.4 * (1 + correctionFactor);
                optionBPercent = 0.5 * (1 + correctionFactor);
                optionCPercent = 0.8 * (1 + correctionFactor);
            }
            // Get total CEC values (handles both least cash outflow and lot-wise)
            let totalCECInclGst = 0;
            if (data.lots && data.lots.length > 0) {
                totalCECInclGst = data.lots.reduce((sum, lot) => sum + (lot.cecEstimateInclGst || 0), 0);
            }
            else {
                totalCECInclGst = data.cecEstimateInclGst || 0;
            }
            // Calculate base values
            const baseOptionA = optionAPercent * totalCECInclGst;
            const baseOptionB = optionBPercent * totalCECInclGst;
            const baseOptionC = optionCPercent * totalCECInclGst;
            // Apply annualization for Service and Works tender types if contract duration > 1 year
            let annualizedOptionA = baseOptionA;
            let annualizedOptionB = baseOptionB;
            let annualizedOptionC = baseOptionC;
            const contractDurationYears = data.contractDurationYears || 1;
            if ((data.tenderType === 'Service' || data.tenderType === 'Works') && contractDurationYears > 1) {
                annualizedOptionA = baseOptionA / contractDurationYears;
                annualizedOptionB = baseOptionB / contractDurationYears;
                annualizedOptionC = baseOptionC / contractDurationYears;
            }
            // Apply MSE relaxation for Service/Works tenders with least cash outflow if enabled
            let finalOptionA = annualizedOptionA;
            let finalOptionB = annualizedOptionB;
            let finalOptionC = annualizedOptionC;
            if ((data.tenderType === 'Service' || data.tenderType === 'Works') && data.evaluationMethodology === 'least cash outflow' && data.mseRelaxation) {
                // Apply 15% relaxation for MSE
                finalOptionA = annualizedOptionA * 0.85;
                finalOptionB = annualizedOptionB * 0.85;
                finalOptionC = annualizedOptionC * 0.85;
            }
            return {
                optionA: {
                    percentage: optionAPercent * 100,
                    value: finalOptionA
                },
                optionB: {
                    percentage: optionBPercent * 100,
                    value: finalOptionB
                },
                optionC: {
                    percentage: optionCPercent * 100,
                    value: finalOptionC
                }
            };
        };
        const formatTurnoverAmount = (amountInCrores) => {
            // Always display in Crores format, rounded to 2 decimal places
            return `Rs. ${Math.round(amountInCrores * 100) / 100} Crore`;
        };
        // Calculate dynamic section numbers based on document structure
        const getSectionNumbers = () => {
            const sections = {
                bqc: 3, // BID QUALIFICATION CRITERIA
                evaluation: 5, // EVALUATION METHODOLOGY  
                emd: 6, // EARNEST MONEY DEPOSIT
                performanceSecurity: 7 // PERFORMANCE SECURITY
            };
            // Adjust section numbers based on document content
            let currentSection = 3; // Start from BQC section
            // BQC section is always 3
            sections.bqc = currentSection;
            currentSection += 1;
            // Evaluation methodology section
            sections.evaluation = currentSection;
            currentSection += 1;
            // EMD section
            sections.emd = currentSection;
            currentSection += 1;
            // Performance Security section (only if applicable)
            if (bqcData.hasPerformanceSecurity) {
                sections.performanceSecurity = currentSection;
                currentSection += 1;
            }
            return sections;
        };
        const sectionNumbers = getSectionNumbers();
        // Format number with Indian comma separation for Crore values
        const formatIndianCurrency = (amountInCrores) => {
            // Convert Crores to actual amount (multiply by 1,00,00,000)
            const actualAmount = amountInCrores * 10000000;
            // Format with Indian number system (lakhs and crores)
            const formattedAmount = actualAmount.toLocaleString('en-IN');
            return `₹ ${formattedAmount}`;
        };
        // Format past performance amount as units (for Goods with least cash outflow)
        const formatPastPerformanceUnits = (units) => {
            return `${units.toLocaleString()} unit${units !== 1 ? 's' : ''}`;
        };
        // Format date to dd/mm/yyyy format
        const formatDate = (date) => {
            let dateObj;
            if (typeof date === 'string') {
                // Handle different input formats
                if (date.includes('-')) {
                    // Handle dd-mm-yyyy or yyyy-mm-dd format
                    const parts = date.split('-');
                    if (parts[0].length === 4) {
                        // yyyy-mm-dd format
                        dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                    }
                    else {
                        // dd-mm-yyyy format
                        dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
                    }
                }
                else {
                    dateObj = new Date(date);
                }
            }
            else {
                dateObj = date;
            }
            // Format as dd/mm/yyyy
            const day = dateObj.getDate().toString().padStart(2, '0');
            const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
            const year = dateObj.getFullYear();
            return `${day}/${month}/${year}`;
        };
        // Get performance security percentage based on tender type
        const getPerformanceSecurityPercentage = (tenderType) => {
            if (tenderType === 'Works') {
                return '10% (Works)';
            }
            else {
                // For Goods and Service
                return '5% (Goods & Service)';
            }
        };
        // Format experience requirements in Crores
        const formatExperienceCurrency = (amount) => {
            return `Rs. ${Math.round(amount * 100) / 100} Crore`;
        };
        // Calculate values based on methodology
        let pastPerformanceNonMSE = 0;
        let pastPerformanceMSE = 0;
        let turnoverAmount = 0;
        let emdAmount = 0;
        let totalCECInclGst = 0;
        let totalCECExclGst = 0;
        let experienceRequirements = null;
        if (bqcData.evaluationMethodology === 'least cash outflow') {
            // least cash outflow methodology - use main CEC values
            const quantitySupplied = bqcData.quantitySupplied || 0;
            // Only calculate past performance for Goods tender type
            if (bqcData.tenderType === 'Goods') {
                pastPerformanceNonMSE = calculatePastPerformance(quantitySupplied, false);
                // Use the specific MSE relaxation field for Past Performance Requirement
                const mseRelaxation = bqcData.pastPerformanceMseRelaxation || false;
                pastPerformanceMSE = calculatePastPerformance(quantitySupplied, mseRelaxation);
            }
            turnoverAmount = calculateTurnover(bqcData);
            emdAmount = calculateEMD(bqcData.cecEstimateInclGst || 0, bqcData.tenderType || 'Goods');
            totalCECInclGst = bqcData.cecEstimateInclGst || 0;
            totalCECExclGst = bqcData.cecEstimateExclGst || 0;
            experienceRequirements = calculateExperienceRequirements(bqcData);
        }
        else {
            // Lot-wise methodology - calculate from lots
            if (bqcData.lots && bqcData.lots.length > 0) {
                totalCECInclGst = bqcData.lots.reduce((sum, lot) => sum + (lot.cecEstimateInclGst || 0), 0);
                totalCECExclGst = bqcData.lots.reduce((sum, lot) => sum + (lot.cecEstimateExclGst || 0), 0);
                // Only calculate past performance for Goods tender type
                if (bqcData.tenderType === 'Goods') {
                    pastPerformanceNonMSE = bqcData.lots.reduce((sum, lot) => sum + calculatePastPerformance(lot.quantitySupplied || 0, false), 0);
                    pastPerformanceMSE = bqcData.lots.reduce((sum, lot) => sum + calculatePastPerformance(lot.quantitySupplied || 0, lot.mseRelaxation || false), 0);
                }
                turnoverAmount = calculateTurnover(bqcData);
                emdAmount = calculateEMD(totalCECInclGst, bqcData.tenderType || 'Goods');
                experienceRequirements = calculateExperienceRequirements(bqcData);
            }
        }
        // Create document with BPCL format using tables
        const doc = new docx_1.Document({
            sections: [{
                    properties: {},
                    children: [
                        // Header Table
                        new docx_1.Table({
                            width: {
                                size: 100,
                                type: docx_1.WidthType.PERCENTAGE,
                            },
                            borders: {
                                top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                            },
                            rows: [
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: `Ref: ${bqcData.refNumber || 'XXXXXX'}`,
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 200 },
                                                }),
                                            ],
                                            width: { size: 50, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: `Date: ${formatDate(new Date())}`,
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.RIGHT,
                                                    spacing: { after: 200 },
                                                }),
                                            ],
                                            width: { size: 50, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                            ],
                        }),
                        // Note To Table
                        new docx_1.Table({
                            width: {
                                size: 100,
                                type: docx_1.WidthType.PERCENTAGE,
                            },
                            borders: {
                                top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                            },
                            rows: [
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "NOTE TO:",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.noteTo || "CHIEF PROCUREMENT OFFICER, CPO (M)",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 80, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                            ],
                        }),
                        // Subject Table
                        new docx_1.Table({
                            width: {
                                size: 100,
                                type: docx_1.WidthType.PERCENTAGE,
                            },
                            borders: {
                                top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                            },
                            rows: [
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "SUBJECT:",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.subject || "APPROVAL OF BID QUALIFICATION CRITERIA AND FLOATING OF OPEN DOMESTIC TENDER.",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 80, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                            ],
                        }),
                        // Add spacing between top section and Preamble
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 400 },
                        }),
                        // Preamble Section Table
                        new docx_1.Table({
                            width: {
                                size: 100,
                                type: docx_1.WidthType.PERCENTAGE,
                            },
                            borders: {
                                top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                            },
                            rows: [
                                // Header row
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "1. PREAMBLE",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.CENTER,
                                                    spacing: { after: 200 },
                                                }),
                                            ],
                                            width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                                            borders: {
                                                top: { style: docx_1.BorderStyle.NONE, size: 0 },
                                                bottom: { style: docx_1.BorderStyle.NONE, size: 0 },
                                                left: { style: docx_1.BorderStyle.NONE, size: 0 },
                                                right: { style: docx_1.BorderStyle.NONE, size: 0 },
                                            },
                                        }),
                                    ],
                                }),
                                // Tender Description
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Tender Description",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.tenderDescription || "N/A",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // PR reference
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "PR reference/ Email reference",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.prReference || "N/A",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // Type of Tender
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Type of Tender",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.tenderType || "Goods",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // CEC estimate incl GST
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "CEC estimate (incl. of GST)/ Date",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: `${formatIndianCurrency(totalCECInclGst || 0)} / ${bqcData.cecDate ? formatDate(bqcData.cecDate) : "N/A"}`,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // CEC estimate excl GST
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "CEC estimate exclusive of GST",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: `${formatIndianCurrency(totalCECExclGst || 0)}`,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // Budget Details
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Budget Details (WBS/ Revex)",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.budgetDetails || "N/A",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // Tender Platform
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Tender Platform",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.tenderPlatform || "GeM",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                            ],
                        }),
                        // Add spacing between sections
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 400 },
                        }),
                        // Brief Scope of Work Section Table
                        new docx_1.Table({
                            width: {
                                size: 100,
                                type: docx_1.WidthType.PERCENTAGE,
                            },
                            borders: {
                                top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                            },
                            rows: [
                                // Header row
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "2. BRIEF SCOPE OF WORK",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.CENTER,
                                                    spacing: { after: 200 },
                                                }),
                                            ],
                                            width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                                            borders: {
                                                top: { style: docx_1.BorderStyle.NONE, size: 0 },
                                                bottom: { style: docx_1.BorderStyle.NONE, size: 0 },
                                                left: { style: docx_1.BorderStyle.NONE, size: 0 },
                                                right: { style: docx_1.BorderStyle.NONE, size: 0 },
                                            },
                                        }),
                                    ],
                                }),
                                // Brief Scope of Work
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Brief Scope of Work / Supply Items",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.scopeOfWork || "N/A",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // Contract Period
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Contract Period /Completion Period",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.contractPeriodText || `${bqcData.contractPeriodMonths || 12} months`,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // Delivery Period
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Delivery Period of the Item",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.deliveryPeriod || "N/A",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // Warranty Period
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Warranty Period",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.warrantyPeriod || "N/A",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // AMC Period - only show if hasAmc and amcValue > 0
                                ...(bqcData.hasAmc && bqcData.amcValue && bqcData.amcValue > 0 ? [new docx_1.TableRow({
                                        children: [
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "AMC/ CAMC/ O&M (No. of Years)",
                                                                bold: true,
                                                                size: 24,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.LEFT,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: `${bqcData.amcPeriod || "N/A"} years`,
                                                                size: 24,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.LEFT,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                        ],
                                    })] : []),
                                // Payment Terms
                                new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: "Payment Terms (if different from standard terms i.e within 30 days)",
                                                            bold: true,
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                        new docx_1.TableCell({
                                            children: [
                                                new docx_1.Paragraph({
                                                    children: [
                                                        new docx_1.TextRun({
                                                            text: bqcData.paymentTerms || "Within 30 days",
                                                            size: 24,
                                                            font: "Arial"
                                                        }),
                                                    ],
                                                    alignment: docx_1.AlignmentType.LEFT,
                                                    spacing: { after: 100 },
                                                }),
                                            ],
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                            ],
                        }),
                        // BQC Section
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `${sectionNumbers.bqc}. BID QUALIFICATION CRITERIA (BQC)`,
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "BPCL would like to qualify vendors for undertaking the above work as indicated in the brief scope. Detailed bid qualification criteria for short listing vendors shall be as follows:",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "3.1 TECHNICAL CRITERIA",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        // Dynamic content based on tender type
                        ...(bqcData.tenderType === 'Goods' ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "3.1.1. For GOODS:",
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Manufacturing Capability:",
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 100 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: `Bidder* should be ${bqcData.manufacturerTypes?.join(' AND/OR ') || 'Original Equipment Manufacturer AND/OR Authorized Channel Partner AND/OR Authorized Agent AND/OR Dealer AND/OR Authorized Distributor'} of the item being tendered.`,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            // Only show Supplying Capacity section for Goods tender type with least cash outflow methodology
                            ...(bqcData.tenderType === 'Goods' && bqcData.evaluationMethodology === 'least cash outflow' ? [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "Supplying Capacity:",
                                            bold: true,
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 100 },
                                }),
                                // Always show Non-MSE (Standard) Requirements
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "Non-MSE (Standard) Requirements:",
                                            bold: true,
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 100 },
                                }),
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: `The bidder shall have experience of having successfully supplied minimum of ${formatPastPerformanceUnits(pastPerformanceNonMSE)} in any 12 continuous months during last 7 years in India or abroad, ending on last day of the month previous to the one in which tender is invited.`,
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 200 },
                                }),
                                // Show MSE (Relaxed) Requirements only when MSE relaxation is enabled
                                ...(bqcData.pastPerformanceMseRelaxation ? [
                                    new docx_1.Paragraph({
                                        children: [
                                            new docx_1.TextRun({
                                                text: "MSE (Relaxed) Requirements:",
                                                bold: true,
                                                size: 24,
                                                font: "Arial"
                                            }),
                                        ],
                                        spacing: { after: 100 },
                                    }),
                                    new docx_1.Paragraph({
                                        children: [
                                            new docx_1.TextRun({
                                                text: `The MSE bidder shall have experience of having successfully supplied minimum of ${formatPastPerformanceUnits(pastPerformanceMSE)} in any 12 continuous months during last 7 years in India or abroad, ending on last day of the month previous to the one in which tender is invited.`,
                                                size: 24,
                                                font: "Arial"
                                            }),
                                        ],
                                        spacing: { after: 200 },
                                    }),
                                    new docx_1.Paragraph({
                                        children: [
                                            new docx_1.TextRun({
                                                text: "For MSE bidders Relaxation of 15% on the supplying capacity shall be given as per Corp. Finance Circular MA.TEC.POL.CON.3A dated 26.10.2020.",
                                                size: 24,
                                                font: "Arial"
                                            }),
                                        ],
                                        spacing: { after: 200 },
                                    }),
                                ] : []),
                            ] : []),
                        ] : []),
                        // Lot-wise Supplying Capacity table for Goods tender type
                        ...(bqcData.tenderType === 'Goods' && bqcData.evaluationMethodology === 'Lot-wise' && bqcData.lots && bqcData.lots.length > 0 ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Supplying Capacity:",
                                        bold: true,
                                        size: 22,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Non-MSE (Standard) Requirements:\nThe bidder should have supplied similar goods in the last Seven (7) years. The quantity supplied should be at least 30% of the total quantity required for each lot as per below table.",
                                        size: 22,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "For MSE bidders, Relaxation of 15% on the supplying capacity shall be given as per Corp. Finance Circular MA.TEC.POL.CON.3A dated 26.10.2020.",
                                        size: 22,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            // Supplying Capacity table with conditional columns
                            (() => {
                                const showNonMse = bqcData.showNonMseCalculations !== false; // Default to true
                                const showMse = bqcData.showMseCalculations !== false; // Default to true
                                // Calculate dynamic widths
                                const remainingWidth = 100 - 10 - 30 - 20; // Total - Sr.No - Description - Quantity
                                const dynamicColumns = (showNonMse ? 1 : 0) + (showMse ? 1 : 0);
                                const dynamicCellWidth = dynamicColumns > 0 ? Math.floor(remainingWidth / dynamicColumns) : 0;
                                return new docx_1.Table({
                                    width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                                    borders: {
                                        top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    },
                                    rows: [
                                        // Header row
                                        new docx_1.TableRow({
                                            children: [
                                                new docx_1.TableCell({
                                                    children: [new docx_1.Paragraph({
                                                            children: [new docx_1.TextRun({ text: "Sr. No.", bold: true, size: 20, font: "Arial" })],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                        })],
                                                    width: { size: 10, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [new docx_1.Paragraph({
                                                            children: [new docx_1.TextRun({ text: "Section / Description", bold: true, size: 20, font: "Arial" })],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                        })],
                                                    width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [new docx_1.Paragraph({
                                                            children: [new docx_1.TextRun({ text: "Quantity Required", bold: true, size: 20, font: "Arial" })],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                        })],
                                                    width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                ...(showNonMse ? [new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: "Non-MSE (30%)", bold: true, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                            })],
                                                        width: { size: dynamicCellWidth, type: docx_1.WidthType.PERCENTAGE },
                                                    })] : []),
                                                ...(showMse ? [new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: "MSE (15%)", bold: true, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                            })],
                                                        width: { size: dynamicCellWidth, type: docx_1.WidthType.PERCENTAGE },
                                                    })] : []),
                                            ],
                                        }),
                                        // Data rows
                                        ...bqcData.lots.map((lot, index) => {
                                            // Parse quantitySupplied - handle both string and number types
                                            const quantityRequired = typeof lot.quantitySupplied === 'string'
                                                ? parseFloat(lot.quantitySupplied) || 0
                                                : (lot.quantitySupplied || 0);
                                            const nonMseRequirement = Math.round(quantityRequired * 0.3);
                                            const mseRequirement = Math.round(quantityRequired * 0.15); // 15% for MSE
                                            console.log(`  Supplying Capacity Table - Lot ${index + 1} (${lot.lotNumber}):`, {
                                                quantitySupplied: lot.quantitySupplied,
                                                quantitySuppliedType: typeof lot.quantitySupplied,
                                                quantityRequired,
                                                nonMseRequirement,
                                                mseRequirement
                                            });
                                            return new docx_1.TableRow({
                                                children: [
                                                    new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: `${index + 1}`, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                            })],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: `${lot.lotNumber || `LOT-${index + 1}`}`, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.LEFT,
                                                            })],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: `${quantityRequired.toLocaleString()}`, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                            })],
                                                    }),
                                                    ...(showNonMse ? [new docx_1.TableCell({
                                                            children: [new docx_1.Paragraph({
                                                                    children: [new docx_1.TextRun({ text: `${nonMseRequirement.toLocaleString()}`, size: 20, font: "Arial" })],
                                                                    alignment: docx_1.AlignmentType.CENTER,
                                                                })],
                                                        })] : []),
                                                    ...(showMse ? [new docx_1.TableCell({
                                                            children: [new docx_1.Paragraph({
                                                                    children: [new docx_1.TextRun({ text: `${mseRequirement.toLocaleString()}`, size: 20, font: "Arial" })],
                                                                    alignment: docx_1.AlignmentType.CENTER,
                                                                })],
                                                        })] : []),
                                                ],
                                            });
                                        }),
                                        // Total row
                                        (() => {
                                            const totalQuantity = bqcData.lots.reduce((sum, lot) => {
                                                const qty = typeof lot.quantitySupplied === 'string'
                                                    ? parseFloat(lot.quantitySupplied) || 0
                                                    : (lot.quantitySupplied || 0);
                                                return sum + qty;
                                            }, 0);
                                            const totalNonMse = Math.round(totalQuantity * 0.3);
                                            const totalMse = Math.round(totalQuantity * 0.15);
                                            console.log(`  Supplying Capacity Table - TOTAL:`, {
                                                totalQuantity,
                                                totalNonMse,
                                                totalMse
                                            });
                                            return new docx_1.TableRow({
                                                children: [
                                                    new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: `${bqcData.lots.length + 1}`, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                            })],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: "TOTAL FOR ALL LOTS", bold: true, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.LEFT,
                                                            })],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [new docx_1.Paragraph({
                                                                children: [new docx_1.TextRun({ text: `${totalQuantity.toLocaleString()}`, size: 20, font: "Arial" })],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                            })],
                                                    }),
                                                    ...(showNonMse ? [new docx_1.TableCell({
                                                            children: [new docx_1.Paragraph({
                                                                    children: [new docx_1.TextRun({ text: `${totalNonMse.toLocaleString()}`, size: 20, font: "Arial" })],
                                                                    alignment: docx_1.AlignmentType.CENTER,
                                                                })],
                                                        })] : []),
                                                    ...(showMse ? [new docx_1.TableCell({
                                                            children: [new docx_1.Paragraph({
                                                                    children: [new docx_1.TextRun({ text: `${totalMse.toLocaleString()}`, size: 20, font: "Arial" })],
                                                                    alignment: docx_1.AlignmentType.CENTER,
                                                                })],
                                                        })] : []),
                                                ],
                                            });
                                        })(),
                                    ],
                                });
                            })(),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Bidder can quote for any one or more than one LOT based on their capability/choice.",
                                        size: 22,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { before: 200, after: 400 },
                            }),
                        ] : []),
                        // Explanatory Note for Past Performance Requirement
                        ...(bqcData.hasPastPerformanceExplanatoryNote && bqcData.pastPerformanceExplanatoryNote ? [
                            new docx_1.Paragraph({
                                children: (0, htmlToWord_js_1.convertHtmlToWordRuns)(bqcData.pastPerformanceExplanatoryNote),
                                spacing: { after: 200 },
                            }),
                        ] : []),
                        // Service/Works content - Exclude Lot-wise as it has its own section
                        ...((bqcData.tenderType === 'Service' || bqcData.tenderType === 'Works') && bqcData.evaluationMethodology !== 'Lot-wise' ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "3.1.2. BQC/PQC for Procurement of Works and Services:",
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            // Experience / Past performance / Technical Capability section - only for non-lot-wise methodology
                            ...(bqcData.evaluationMethodology !== 'Lot-wise' ? [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "Experience / Past performance / Technical Capability:",
                                            bold: true,
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 100 },
                                }),
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "The bidder should have experience of having successfully completed similar works during last 7 years ending last day of month previous to the one in which tender is floated should be either of the following: -",
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 200 },
                                }),
                                // Experience Requirements - Always show for Works/Service with least cash outflow
                                // Show for Lot-wise OR for least cash outflow (regardless of MSE relaxation)
                                ...((bqcData.tenderType === 'Service' || bqcData.tenderType === 'Works') &&
                                    (bqcData.evaluationMethodology === 'Lot-wise' || bqcData.evaluationMethodology === 'least cash outflow') ? [
                                    new docx_1.Paragraph({
                                        children: [
                                            new docx_1.TextRun({
                                                text: "Experience Requirements:",
                                                bold: true,
                                                size: 24,
                                                font: "Arial"
                                            }),
                                        ],
                                        spacing: { after: 200 },
                                    }),
                                    // Show standard requirements first (only for least cash outflow when MSE is enabled)
                                    ...(bqcData.evaluationMethodology === 'least cash outflow' && bqcData.mseRelaxation ? [
                                        // Standard Requirements (reverse of MSE adjustment: divide by 0.85)
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "Standard Requirements:",
                                                    bold: true,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `Three similar completed works each costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionA.value / 0.85 : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "or",
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `Two similar completed works each costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionB.value / 0.85 : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "or",
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `One similar completed work costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionC.value / 0.85 : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 200 },
                                        }),
                                    ] : []),
                                    // Show standard requirements for least cash outflow when MSE relaxation is NOT enabled
                                    ...(bqcData.evaluationMethodology === 'least cash outflow' && !bqcData.mseRelaxation ? [
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `Three similar completed works each costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionA.value : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "or",
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `Two similar completed works each costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionB.value : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "or",
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `One similar completed work costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionC.value : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 200 },
                                        }),
                                    ] : []),
                                    // Show MSE-specific content for least cash outflow when MSE relaxation is enabled
                                    ...(bqcData.evaluationMethodology === 'least cash outflow' && bqcData.mseRelaxation ? [
                                        // MSE Relaxed Requirements
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "MSE Relaxed Requirements (15% reduction):",
                                                    bold: true,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `Three similar completed works each costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionA.value : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "or",
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `Two similar completed works each costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionB.value : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: "or",
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 100 },
                                        }),
                                        new docx_1.Paragraph({
                                            children: [
                                                new docx_1.TextRun({
                                                    text: `One similar completed work costing not less than ${formatExperienceCurrency(experienceRequirements ? experienceRequirements.optionC.value : 0)}.`,
                                                    size: 24,
                                                    font: "Arial"
                                                }),
                                            ],
                                            spacing: { after: 200 },
                                        }),
                                    ] : []),
                                    new docx_1.Paragraph({
                                        children: [
                                            new docx_1.TextRun({
                                                text: `Definition of the similar work should be considered as following: ${bqcData.similarWorkDefinition || "N/A"}`,
                                                size: 24,
                                                font: "Arial"
                                            }),
                                        ],
                                        spacing: { after: 200 },
                                    }),
                                ] : []),
                            ] : []),
                        ] : []),
                        // Explanatory Note for Experience Requirements
                        ...(bqcData.hasExperienceExplanatoryNote && bqcData.experienceExplanatoryNote ? [
                            new docx_1.Paragraph({
                                children: (0, htmlToWord_js_1.convertHtmlToWordRuns)(bqcData.experienceExplanatoryNote),
                                spacing: { after: 200 },
                            }),
                        ] : []),
                        // Lot-wise Technical Criteria Section - Only for Lot-wise methodology
                        ...(bqcData.evaluationMethodology === 'Lot-wise' && bqcData.lots && bqcData.lots.length > 0 && (bqcData.tenderType === 'Service' || bqcData.tenderType === 'Works') ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "3.1.2. BQC/PQC for Procurement of Works and Services",
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "3.1.1 PROVEN TRACK RECORD",
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "The bidder shall have experience of having successfully executed similar works in the last Seven (7) years in any Oil & Gas Industry in India. The Value (Rs) of the similar work/s executed (proof of execution to be submitted) should be as follows:",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Definition of \"similar work\": Similar works shall be considered as – \"Bidder should have executed the job of Pipeline Works for Hydrocarbons/Petrochemicals/ Fertilizers/Chemicals/ Fire Fighting system, with or without associated works.\"",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            // Standard Requirements Table (always shown)
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Standard Requirements (Non-MSE)",
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 100 },
                            }),
                            new docx_1.Table({
                                width: {
                                    size: 100,
                                    type: docx_1.WidthType.PERCENTAGE,
                                },
                                borders: {
                                    top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                },
                                rows: [
                                    // Header row
                                    new docx_1.TableRow({
                                        children: [
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Sr. No.",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 15, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Section / Description",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 25, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "One similar work of total value not less than (Rs. in Lakhs)",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Two similar works EACH of value not less than (Rs. in Lakhs)",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Three similar works EACH of value not less than (Rs. in Lakhs)",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                        ],
                                    }),
                                    // Data rows for each lot
                                    ...bqcData.lots.map((lot, index) => {
                                        console.log(`🔍 DEBUG: Processing lot ${index + 1} (${lot.lotNumber}) for Technical Criteria - Similar Works`);
                                        console.log(`  Raw cecEstimateInclGst:`, lot.cecEstimateInclGst);
                                        console.log(`  Type:`, typeof lot.cecEstimateInclGst);
                                        console.log(`  Is number?:`, typeof lot.cecEstimateInclGst === 'number');
                                        console.log(`  Is undefined?:`, lot.cecEstimateInclGst === undefined);
                                        console.log(`  Is null?:`, lot.cecEstimateInclGst === null);
                                        console.log(`  Is 0?:`, lot.cecEstimateInclGst === 0);
                                        const baseAmount = lot.cecEstimateInclGst || 0;
                                        console.log(`  Parsed baseAmount:`, baseAmount);
                                        if (baseAmount === 0) {
                                            console.error(`❌ ERROR: Lot ${index + 1} has baseAmount = 0. This will cause all calculations to be 0.`);
                                            console.error(`  Original value was:`, lot.cecEstimateInclGst);
                                        }
                                        // Parse contract period from text or use numeric value
                                        let contractMonths = lot.contractPeriodMonths || 12;
                                        if (lot.contractPeriodText) {
                                            const textMatch = lot.contractPeriodText.match(/(\d+)/);
                                            if (textMatch) {
                                                contractMonths = parseInt(textMatch[1]);
                                                // Handle years conversion
                                                if (lot.contractPeriodText.toLowerCase().includes('year')) {
                                                    contractMonths = contractMonths * 12;
                                                }
                                            }
                                        }
                                        const contractYears = contractMonths / 12;
                                        const annualizedAmount = contractYears > 1 ? baseAmount / contractYears : baseAmount;
                                        console.log(`  contractYears:`, contractYears);
                                        console.log(`  annualizedAmount:`, annualizedAmount);
                                        // Convert to Lakhs for display (1 Crore = 100 Lakhs)
                                        // annualizedAmount is in Crores, so multiply by 100 to get Lakhs
                                        const annualizedAmountInLakhs = annualizedAmount * 100;
                                        // Standard values (no MSE reduction) - Calculate in Lakhs
                                        const optionA = annualizedAmountInLakhs * 0.8; // 80% - One work
                                        const optionB = annualizedAmountInLakhs * 0.5; // 50% - Two works each
                                        const optionC = annualizedAmountInLakhs * 0.4; // 40% - Three works each
                                        console.log(`  Final values - optionA: ${optionA} Lakhs, optionB: ${optionB} Lakhs, optionC: ${optionC} Lakhs`);
                                        console.log(`  RENDERING AS: optionA=${optionA.toFixed(2)}, optionB=${optionB.toFixed(2)}, optionC=${optionC.toFixed(2)}`);
                                        return new docx_1.TableRow({
                                            children: [
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${index + 1}`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: lot.lotNumber || `LOT-${index + 1} (${lot.description || 'Region'})`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${(Math.round(optionA * 100) / 100).toFixed(2)}`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${(Math.round(optionB * 100) / 100).toFixed(2)}`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${(Math.round(optionC * 100) / 100).toFixed(2)}`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        });
                                    }),
                                ],
                            }),
                            // MSE Requirements Table - Only shown when MSE relaxation is enabled
                            ...(bqcData.provenTrackRecordMseRelaxation ? [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "MSE Requirements (15% Reduction)",
                                            bold: true,
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { before: 200, after: 100 },
                                }),
                                new docx_1.Table({
                                    width: {
                                        size: 100,
                                        type: docx_1.WidthType.PERCENTAGE,
                                    },
                                    borders: {
                                        top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    },
                                    rows: [
                                        // Header row
                                        new docx_1.TableRow({
                                            children: [
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "Sr. No.",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 15, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "Section / Description",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 25, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "One similar work of total value not less than (Rs. in Lakhs)",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "Two similar works EACH of value not less than (Rs. in Lakhs)",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "Three similar works EACH of value not less than (Rs. in Lakhs)",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 20, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                            ],
                                        }),
                                        // Data rows for each lot with MSE reduction
                                        ...bqcData.lots.map((lot, index) => {
                                            const baseAmount = lot.cecEstimateInclGst || 0;
                                            // Parse contract period from text or use numeric value
                                            let contractMonths = lot.contractPeriodMonths || 12;
                                            if (lot.contractPeriodText) {
                                                const textMatch = lot.contractPeriodText.match(/(\d+)/);
                                                if (textMatch) {
                                                    contractMonths = parseInt(textMatch[1]);
                                                    // Handle years conversion
                                                    if (lot.contractPeriodText.toLowerCase().includes('year')) {
                                                        contractMonths = contractMonths * 12;
                                                    }
                                                }
                                            }
                                            const contractYears = contractMonths / 12;
                                            const annualizedAmount = contractYears > 1 ? baseAmount / contractYears : baseAmount;
                                            // Convert to Lakhs for display (1 Crore = 100 Lakhs)
                                            const annualizedAmountInLakhs = annualizedAmount * 100;
                                            // MSE values (with 15% reduction) - Apply reduction to Lakhs
                                            const mseAnnualizedAmountInLakhs = annualizedAmountInLakhs * 0.85;
                                            const mseOptionA = mseAnnualizedAmountInLakhs * 0.8; // 80% - One work
                                            const mseOptionB = mseAnnualizedAmountInLakhs * 0.5; // 50% - Two works each
                                            const mseOptionC = mseAnnualizedAmountInLakhs * 0.4; // 40% - Three works each
                                            console.log(`  MSE Table - Lot ${index + 1}: annualizedAmountInLakhs=${annualizedAmountInLakhs}, mseAnnualizedAmountInLakhs=${mseAnnualizedAmountInLakhs}, mseOptionA=${mseOptionA}, mseOptionB=${mseOptionB}, mseOptionC=${mseOptionC}`);
                                            return new docx_1.TableRow({
                                                children: [
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: `${index + 1}`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: lot.lotNumber || `LOT-${index + 1} (${lot.description || 'Region'})`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: `${(Math.round(mseOptionA * 100) / 100).toFixed(2)}`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: `${(Math.round(mseOptionB * 100) / 100).toFixed(2)}`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: `${(Math.round(mseOptionC * 100) / 100).toFixed(2)}`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                ],
                                            });
                                        }),
                                    ],
                                }),
                            ] : []),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Bidder can quote for any one or more than one LOT based on their capability/choice.",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 100 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Bidder can quote for any one or more than one LOT based on their capability/choice.",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 100 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Note: If the Bidder quotes for more than one LOT, the similar works criteria should not be less than the cumulative amount applicable for the LOTs quoted.",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 400 },
                            }),
                        ] : []),
                        // Additional Details
                        ...(bqcData.additionalDetails ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "ADDITIONAL DETAILS",
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: bqcData.additionalDetails,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 400 },
                            }),
                        ] : []),
                        // Explanatory Note for Additional Details
                        ...(bqcData.hasAdditionalExplanatoryNote && bqcData.additionalExplanatoryNote ? [
                            new docx_1.Paragraph({
                                children: (0, htmlToWord_js_1.convertHtmlToWordRuns)(bqcData.additionalExplanatoryNote),
                                spacing: { after: 200 },
                            }),
                        ] : []),
                        // Financial Criteria
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "3.2 FINANCIAL CRITERIA",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "3.2.1 ANNUAL TURNOVER",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 100 },
                        }),
                        ...(bqcData.evaluationMethodology === 'Lot-wise' && bqcData.lots && bqcData.lots.length > 0 ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "The bidder should have achieved a minimum Average Annual financial turnover as per below table (LOT-WISE).as per Audited Balance sheet and P&L Statement in the last three* accounting years prior to due date of bid submission.",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                        ] : []),
                        // Lot-wise Financial Criteria table
                        ...(bqcData.evaluationMethodology === 'Lot-wise' && bqcData.lots && bqcData.lots.length > 0 ? [
                            new docx_1.Table({
                                width: {
                                    size: 100,
                                    type: docx_1.WidthType.PERCENTAGE,
                                },
                                borders: {
                                    top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                },
                                rows: [
                                    // Header row
                                    new docx_1.TableRow({
                                        children: [
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Sr. No.",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 15, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Section / Description",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 35, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Annualized Estimated Value (Rs. In Lakhs)",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 25, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "Average Annual Turnover (Rs. In Lakhs)",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                                width: { size: 25, type: docx_1.WidthType.PERCENTAGE },
                                            }),
                                        ],
                                    }),
                                    // Data rows for each lot
                                    ...bqcData.lots.map((lot, index) => {
                                        const baseAmount = lot.cecEstimateInclGst || 0;
                                        // Parse contract period from text or use numeric value
                                        let contractMonths = lot.contractPeriodMonths || 12;
                                        if (lot.contractPeriodText) {
                                            const textMatch = lot.contractPeriodText.match(/(\d+)/);
                                            if (textMatch) {
                                                contractMonths = parseInt(textMatch[1]);
                                                // Handle years conversion
                                                if (lot.contractPeriodText.toLowerCase().includes('year')) {
                                                    contractMonths = contractMonths * 12;
                                                }
                                            }
                                        }
                                        const contractYears = contractMonths / 12;
                                        const annualizedAmount = contractYears > 1 ? baseAmount / contractYears : baseAmount;
                                        // Convert to Lakhs for display (1 Crore = 100 Lakhs)
                                        const annualizedAmountInLakhs = annualizedAmount * 100;
                                        const turnoverRequirement = annualizedAmountInLakhs * 0.3; // 30% of annualized amount in Lakhs
                                        console.log(`  Annual Turnover Table - Lot ${index + 1}: annualizedAmount=${annualizedAmount} Cr, annualizedAmountInLakhs=${annualizedAmountInLakhs}, turnoverRequirement=${turnoverRequirement}`);
                                        return new docx_1.TableRow({
                                            children: [
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${index + 1}`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: lot.lotNumber || `LOT-${index + 1} (${lot.description || 'Region'})`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${(Math.round(annualizedAmountInLakhs * 100) / 100).toFixed(2)}`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${(Math.round(turnoverRequirement * 100) / 100).toFixed(2)}`,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        });
                                    }),
                                    // Totals row
                                    new docx_1.TableRow({
                                        children: [
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: `${bqcData.lots.length + 1}`,
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: "TOTAL FOR ALL LOTS",
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: `${((bqcData.lots?.reduce((total, lot) => {
                                                                    const baseAmount = lot.cecEstimateInclGst || 0;
                                                                    let contractMonths = lot.contractPeriodMonths || 12;
                                                                    if (lot.contractPeriodText) {
                                                                        const textMatch = lot.contractPeriodText.match(/(\d+)/);
                                                                        if (textMatch) {
                                                                            contractMonths = parseInt(textMatch[1]);
                                                                            if (lot.contractPeriodText.toLowerCase().includes('year')) {
                                                                                contractMonths = contractMonths * 12;
                                                                            }
                                                                        }
                                                                    }
                                                                    const contractYears = contractMonths / 12;
                                                                    const annualizedAmount = contractYears > 1 ? baseAmount / contractYears : baseAmount;
                                                                    // Convert to Lakhs (multiply by 100)
                                                                    return total + (annualizedAmount * 100);
                                                                }, 0) || 0)).toFixed(2)}`,
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                            }),
                                            new docx_1.TableCell({
                                                children: [
                                                    new docx_1.Paragraph({
                                                        children: [
                                                            new docx_1.TextRun({
                                                                text: `${((bqcData.lots?.reduce((total, lot) => {
                                                                    const baseAmount = lot.cecEstimateInclGst || 0;
                                                                    let contractMonths = lot.contractPeriodMonths || 12;
                                                                    if (lot.contractPeriodText) {
                                                                        const textMatch = lot.contractPeriodText.match(/(\d+)/);
                                                                        if (textMatch) {
                                                                            contractMonths = parseInt(textMatch[1]);
                                                                            if (lot.contractPeriodText.toLowerCase().includes('year')) {
                                                                                contractMonths = contractMonths * 12;
                                                                            }
                                                                        }
                                                                    }
                                                                    const contractYears = contractMonths / 12;
                                                                    const annualizedAmount = contractYears > 1 ? baseAmount / contractYears : baseAmount;
                                                                    // Convert to Lakhs (multiply by 100) then calculate 30%
                                                                    const annualizedAmountInLakhs = annualizedAmount * 100;
                                                                    const turnoverRequirement = annualizedAmountInLakhs * 0.3;
                                                                    return total + turnoverRequirement;
                                                                }, 0) || 0)).toFixed(2)}`,
                                                                bold: true,
                                                                size: 20,
                                                                font: "Arial"
                                                            }),
                                                        ],
                                                        alignment: docx_1.AlignmentType.CENTER,
                                                        spacing: { after: 100 },
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Bidder can quote for any one or more than one LOT based on their capability/choice.",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 100 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Note: If the Bidder quotes for more than one LOT, the average value of Turnover should not be less than the cumulative amount applicable for the LOTs quoted.",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                        ] : [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: `The average annual turnover of the Bidder for last three audited accounting years shall be equal to or more than ${formatTurnoverAmount(turnoverAmount || 0)}.`,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                        ]),
                        // 3.2.2 NET WORTH
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "3.2.2 NET WORTH",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 100 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "The bidder should have positive net worth as per the latest audited financial statement.",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        // Explanatory Note for Financial Criteria
                        ...(bqcData.hasFinancialExplanatoryNote && bqcData.financialExplanatoryNote ? [
                            new docx_1.Paragraph({
                                children: (0, htmlToWord_js_1.convertHtmlToWordRuns)(bqcData.financialExplanatoryNote),
                                spacing: { after: 200 },
                            }),
                        ] : []),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "3.3 BIDS MAY BE SUBMITTED BY",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "3.3.1 An entity (domestic bidder) should have completed 3 financial years of existence as on original due date of tender since date of commencement of business and shall fulfil each BQC eligibility criteria as mentioned above.",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "3.3.2 JV/Consortium bids will not be accepted (i.e. Qualification on the strength of the JV Partners/Consortium Members /Subsidiaries / Group members will not be accepted)",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 400 },
                        }),
                        // Escalation Clause - Only for non-E&P SERVICES groups
                        ...(bqcData.groupName !== '4 - E&P SERVICES' ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "4. ESCALATION/ DE-ESCALATION CLAUSE: Buyer to take approval of the relevant clause, if applicable.",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: bqcData.escalationClause || "N/A",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 400 },
                            })
                        ] : []),
                        // Evaluation Methodology
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `${sectionNumbers.evaluation}. EVALUATION METHODOLOGY`,
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "The tender will be invited through Open tender (Domestic) as two-part bid. The bid qualification evaluation of the received bids will be done as per the above bid qualification criteria and the technical bid of the shortlisted bidders will be evaluated subsequently. The price bids of the bidders who qualify BQC criteria & meet Technical / Commercial requirements of the tender will only be opened and evaluated.",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `The Commercial Evaluation shall be done on ${Array.isArray(bqcData.commercialEvaluationMethod) && bqcData.commercialEvaluationMethod.length > 0 ? bqcData.commercialEvaluationMethod.join(', ') : 'Overall Lowest Basis'}.`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: bqcData.tenderType === 'Works'
                                        ? "The order will be placed based on above methodology AND Purchase preference based on PPP-MII Policy."
                                        : "The order will be placed based on above methodology AND Purchase preference based on MSE/ PPP-MII Policy.",
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `The subject job is ${bqcData.divisibility || 'Non-Divisible'}.`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 400 },
                        }),
                        // EMD - Only show if hasEMDPreview is checked
                        ...(bqcData.hasEMDPreview ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: `${sectionNumbers.emd}. EARNEST MONEY DEPOSIT (EMD)`,
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            // EMD Content - Lot-wise or Single amount
                            ...(bqcData.evaluationMethodology === 'Lot-wise' && bqcData.lots && bqcData.lots.length > 0 ? [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "Bidders are required to provide Earnest Money Deposit as per below table (LOT-WISE):",
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 200 },
                                }),
                                // Lot-wise EMD table
                                new docx_1.Table({
                                    width: {
                                        size: 100,
                                        type: docx_1.WidthType.PERCENTAGE,
                                    },
                                    borders: {
                                        top: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        right: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                    },
                                    rows: [
                                        // Header row
                                        new docx_1.TableRow({
                                            children: [
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "Sr. No.",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 15, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "Section / Description",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 35, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "CEC Estimate (Rs. In Lakhs)",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 25, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "EMD Amount (Rs. In Lakhs)",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                    width: { size: 25, type: docx_1.WidthType.PERCENTAGE },
                                                }),
                                            ],
                                        }),
                                        // Data rows for each lot
                                        ...bqcData.lots.map((lot, index) => {
                                            const lotEMD = calculateEMD(lot.cecEstimateInclGst || 0, bqcData.tenderType || 'Goods');
                                            // Convert CEC to Lakhs for display (1 Crore = 100 Lakhs)
                                            const lotCECInLakhs = (lot.cecEstimateInclGst || 0) * 100;
                                            console.log(`  EMD Table - Lot ${index + 1}: CEC=${lot.cecEstimateInclGst} Cr, CECInLakhs=${lotCECInLakhs}, EMD=${lotEMD}`);
                                            return new docx_1.TableRow({
                                                children: [
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: `${index + 1}`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: lot.lotNumber || `LOT-${index + 1} (${lot.description || 'Region'})`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: `${(Math.round(lotCECInLakhs * 100) / 100).toFixed(2)}`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                    new docx_1.TableCell({
                                                        children: [
                                                            new docx_1.Paragraph({
                                                                children: [
                                                                    new docx_1.TextRun({
                                                                        text: `${Math.round(lotEMD * 10) / 10}`,
                                                                        size: 20,
                                                                        font: "Arial"
                                                                    }),
                                                                ],
                                                                alignment: docx_1.AlignmentType.CENTER,
                                                                spacing: { after: 100 },
                                                            }),
                                                        ],
                                                    }),
                                                ],
                                            });
                                        }),
                                        // Totals row
                                        new docx_1.TableRow({
                                            children: [
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${bqcData.lots.length + 1}`,
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: "TOTAL FOR ALL LOTS",
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${((bqcData.lots?.reduce((total, lot) => {
                                                                        // Convert CEC to Lakhs (multiply by 100)
                                                                        return total + ((lot.cecEstimateInclGst || 0) * 100);
                                                                    }, 0) || 0)).toFixed(2)}`,
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                                new docx_1.TableCell({
                                                    children: [
                                                        new docx_1.Paragraph({
                                                            children: [
                                                                new docx_1.TextRun({
                                                                    text: `${Math.round((bqcData.lots?.reduce((total, lot) => total + calculateEMD(lot.cecEstimateInclGst || 0, bqcData.tenderType || 'Goods'), 0) || 0) * 10) / 10}`,
                                                                    bold: true,
                                                                    size: 20,
                                                                    font: "Arial"
                                                                }),
                                                            ],
                                                            alignment: docx_1.AlignmentType.CENTER,
                                                            spacing: { after: 100 },
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                    ],
                                }),
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "Bidder can quote for any one or more than one LOT based on their capability/choice.",
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 100 },
                                }),
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: "Note: If the Bidder quotes for more than one LOT, the EMD amount should not be less than the cumulative amount applicable for the LOTs quoted.",
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 200 },
                                }),
                            ] : [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: `Bidders are required to provide Earnest Money Deposit equivalent to Rs. ${Math.round((emdAmount || 0) * 100) / 100} Lacs for the tender.`,
                                            size: 24,
                                            font: "Arial"
                                        }),
                                    ],
                                    spacing: { after: 200 },
                                }),
                            ]),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "EMD exemption shall be as per General Terms & Conditions of GeM (applicable for GeM tenders)/ MSE policy",
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 400 },
                            }),
                            // Explanatory Note for EMD
                            ...(bqcData.hasEMDExplanatoryNote && bqcData.emdExplanatoryNote ? [
                                new docx_1.Paragraph({
                                    children: (0, htmlToWord_js_1.convertHtmlToWordRuns)(bqcData.emdExplanatoryNote),
                                    spacing: { after: 200 },
                                }),
                            ] : []),
                        ] : []),
                        // Performance Security - Only if enabled
                        ...(bqcData.hasPerformanceSecurity ? [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: `${sectionNumbers.performanceSecurity}. Performance Security (if at variance with the ITB clause):`,
                                        bold: true,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 200 },
                            }),
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: `Performance Security % other than ${getPerformanceSecurityPercentage(bqcData.tenderType || 'Goods')} to be mentioned, approved by the competent authority: ${bqcData.performanceSecurity || 'Standard'}`,
                                        size: 24,
                                        font: "Arial"
                                    }),
                                ],
                                spacing: { after: 400 },
                            })
                        ] : []),
                        // Approval Required
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "8. APPROVAL REQUIRED",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `In view of above, approval is requested for the supply of ${bqcData.tenderDescription || 'the tender'}/job ${bqcData.tenderDescription || 'the tender'} for`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `Bid Qualification Criteria as per Sr. No. ${sectionNumbers.bqc}, as per Clause 13.8 of Guidelines for procurement of Goods and Contract Services.`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 100 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `Inviting bids (two-part bid) through a Domestic Open Tender and adopting evaluation methodology as per Sr. No. ${sectionNumbers.evaluation} above.`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 100 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `Earnest Money Deposit as per Sr. No. ${sectionNumbers.emd} above.${bqcData.hasPerformanceSecurity ? '/ Performance Security as per Sr. No. ' + sectionNumbers.performanceSecurity + ' (if applicable)' : ''}`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            spacing: { after: 400 },
                        }),
                        // Approval Section - Centered
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "Proposed by",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `${bqcData.proposedBy || "XXXXX"}${bqcData.proposedByDesignation ? ', ' + bqcData.proposedByDesignation : ', Procurement Manager (CPO Mktg.)'}`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 400 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "Recommended by",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `${bqcData.recommendedBy || "XXXXXX"}${bqcData.recommendedByDesignation ? ', ' + bqcData.recommendedByDesignation : ', Procurement Leader (CPO Mktg.)'}`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 400 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "Concurred by",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `${bqcData.concurredBy || "Rajesh J."}${bqcData.concurredByDesignation ? ', ' + bqcData.concurredByDesignation : ', General Manager Finance (CPO Marketing)'}`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 400 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "Approved by",
                                    bold: true,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 200 },
                        }),
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: `${bqcData.approvedBy || "Kani Amudhan N."}${bqcData.approvedByDesignation ? ', ' + bqcData.approvedByDesignation : ', Chief Procurement Officer (CPO Marketing)'}`,
                                    size: 24,
                                    font: "Arial"
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                            spacing: { after: 200 },
                        }),
                    ],
                }],
        });
        const buffer = await docx_1.Packer.toBuffer(doc);
        // Set response headers for file download
        const filename = `BQC_${bqcData.refNumber || 'document'}_${formatDate(new Date()).replace(/\//g, '-')}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Length', buffer.length);
        // Send the buffer
        res.send(buffer);
    }
    catch (error) {
        console.error('Generate document error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to generate document'
        });
    }
});
exports.default = router;
