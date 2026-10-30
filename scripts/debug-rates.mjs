import fs from 'node:fs';
import { Workbook } from './src/lib/engine.js';
import { summarize } from './src/lib/model.js';
import { computeUnitWiseBreakdown } from './src/components/quotation/OfficialQuotationDocument.js';

const rawWb = JSON.parse(fs.readFileSync('data/workbook.json', 'utf8'));
const SERVER_CONSTANTS = { CONST_STEEL_FACTOR: 127388, CONST_SURFACE_FACTOR: 645.16 };
const wb = new Workbook(rawWb, SERVER_CONSTANTS);

console.log('Testing System A & B rates...');
