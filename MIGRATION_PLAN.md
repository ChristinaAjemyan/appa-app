# Calculator To Relational PostgreSQL Migration Plan

## Objective

Move the calculator from arbitrary JSONB key/value persistence to a relational PostgreSQL design while preserving historical calculations and supporting the existing legacy policy data.

The target architecture will use one canonical policy ledger for both calculator data and legacy `insurance_policies` records. Frequently queried business data will be normalized. Complex configuration fragments may remain in versioned JSONB during the first phase.

## Current State

The calculator currently stores business records through `calcStorage`, backed by `calc_result_storage`:

- `month:<YYYY-MM>`: imported OSAGO policies and voluntary policies
- `officePol:<YYYY-MM>`: office-sale policies
- `monthSnapshot:<YYYY-MM>`: rates, voluntary rates, exceptions, agent directory, and manager configuration
- `cashBook:<month>` and `cashBook_mro:<month>`: daily cash-book state
- `officeExpenses`: month-keyed expense rows
- `renewalWork:<month>` and `renewalResult:<type>:<month>`: renewal workflow and cached results
- `tasks`, `bookmarks:<user>`, `reminders`, `auditLog`, and `amexTopups`
- `agentDirectory`, `appSettings`, `ratesConfig`, `volRates`, `exceptionsConfig`, `managerConfig`, and `mreoConfig`

The existing backend has separate models for agents, companies, regions, percentages, users, insurance policies, and JSONB calculator storage. These models currently have little or no foreign-key relationship between them.

## Design Decisions

- **Canonical ledger:** calculator and legacy policies will be migrated into one relational `policies` table.
- **Configuration:** use a hybrid design. Normalize rates, rules, tiers, and frequently queried values; retain only genuinely complex nested fragments in versioned JSONB.
- **Historical correctness:** saved and locked months must be reproducible even after current rates or agent mappings change.
- **Rollout:** use backfill, dual-write, dual-read comparison, pilot cutover, and delayed cleanup.
- **Client storage:** retain browser `localStorage` only for UI preferences such as layout and sort settings.
- **Compatibility:** keep the JSONB store as a temporary read-only archive during migration; do not use it as the long-term source of truth.

## Phase 1: Define The Canonical Data Contract

Before schema migration, document the exact meaning of each calculator object.

### Policy fields

The canonical policy contract should include:

- database ID and legacy/source ID
- source: `agent_import`, `office_sale`, or `legacy`
- product type: `osago` or `voluntary`
- accounting month
- sale date, policy start date, and policy end date
- company
- agent
- policy number
- insured name, phone, passport number, bank account, and email
- car make, car plate, BM class, region, horsepower, and term
- voluntary product name
- status, comment, and source file/import batch
- amount and discount
- payment status, payment type, paid amount, paid date, and payment source
- calculated office, agent, manager, and operator rates
- calculated commission amounts and profit
- calculation version and matched rule/configuration references

### Identity and duplicate rules

Define business rules before backfill:

- client-generated `_id` values must not become database primary keys
- preserve original IDs in source metadata
- detect possible duplicates by policy number, registration plate, insured name, company, and date
- do not automatically delete duplicate candidates until they are reviewed
- preserve whether a row came from an agent import or office sale

## Phase 2: Target Relational Schema

### Reference data

Evolve existing reference tables into canonical relational data:

- `companies`
  - `id`
  - normalized name
  - display name
  - default office rate
  - default agent rate
  - active flag
  - timestamps

- `agents`
  - `id`
  - first name
  - surname
  - internal code
  - active flag
  - timestamps

- `agent_company_codes`
  - `agent_id`
  - `company_id`
  - insurer-specific agent code
  - effective dates
  - unique constraint on company and code

- `regions`
  - retain `region_code` and name
  - add normalized lookup support and foreign-key references

The existing insurer-specific code columns on `agents` should be migrated into `agent_company_codes` rather than extended further.

### Policies

Create a canonical `policies` table with:

- generated primary key
- source and legacy identifiers
- accounting month foreign key
- product type
- company and agent foreign keys
- policy number
- customer/contact fields
- policy dates
- OSAGO vehicle and rating fields
- voluntary product field
- amount, discount, and currency
- policy status and comments
- payment state
- source/import metadata
- timestamps and optional soft-delete fields

Use `NUMERIC`, not `REAL`, for money and percentage values.

### Import provenance

Add:

- `policy_import_batches`
  - source filename
  - source type
  - imported by
  - imported timestamp
  - row counts
  - status
  - idempotency key

- `policy_import_rows`
  - batch ID
  - source row number
  - raw row JSONB
  - normalized policy ID
  - validation status
  - rejection reason

This preserves malformed input and makes repeated imports safe.

### Commission results

Use either a `policy_commissions` table or equivalent child records for:

- office commission
- agent commission
- manager commission
- operator commission
- applied percentage
- calculated amount
- matched rate/rule ID
- calculation version
- historical snapshot reference

Keep policy-level totals for fast reporting, but retain component-level results for auditability.

### Monthly periods and snapshots

Add:

- `monthly_periods`
  - month key
  - status: `open` or `locked`
  - locked by
  - locked at
  - unlock metadata

- `monthly_configuration_snapshots`
  - period ID
  - normalized rate-set references
  - agent/company code snapshot reference
  - manager/operator configuration
  - original snapshot JSONB
  - schema version
  - checksum
  - created by and created at

When a month is saved or locked, store the exact effective configuration used for historical calculations.

### Commission configuration

Normalize the calculator's configuration into:

- `commission_rate_sets`
- `commission_rates`
- `voluntary_product_rates`
- `commission_rules`
- `commission_rule_conditions`
- `manager_rate_sets`
- `manager_tiers`
- `operator_rates`
- `operator_overrides`
- `mreo_configurations`

The rule model must support company, region, BM, horsepower, term, brand, agent inclusion/exclusion, effective dates, and explicit priority.

Retain JSONB only for deeply nested structures that are not yet queried independently. Every JSONB configuration record must have a schema version and effective period.

### Expenses and cash

Replace month-keyed JSON with:

- `expense_categories`
- `office_expenses`
- `cash_books`
- `cash_book_days`
- `cash_book_entries`

Cash records should support main and MREO modes, daily close/reopen state, timestamps, responsible user, payment source, and links to policies/payment records.

### Supporting workflows

Normalize the remaining workflows:

- `renewal_cases`
- `renewal_calls`
- `renewal_status_history`
- `tasks`
- `task_comments`
- `bookmarks`
- `reminders`
- `audit_entries`
- `amex_topups`

Link these records to users, agents, policies, and monthly periods with foreign keys wherever applicable. Renewal calculation caches may remain JSONB temporarily.

### Users and permissions

Move calculator employees and tab access out of `appSettings` into relational records:

- users or employee profiles
- roles
- permissions
- employee role assignments
- tab/access assignments

Do not migrate plaintext PINs. Preserve the existing PBKDF2 format or replace it with a server-side password hashing scheme.

## Phase 3: Database And ORM Foundation

- Replace production use of `sequelize.sync({ alter: true })` with explicit Sequelize migrations.
- Add model associations and foreign-key constraints.
- Standardize table and column naming using plural snake_case for new tables.
- Add timestamps consistently.
- Add checks for valid product types, month formats, amount ranges, and payment states.
- Add indexes for:
  - accounting month, source, and product type
  - company and agent
  - policy number
  - registration plate
  - insured name and phone
  - policy end date
  - payment state
  - cash-book date
  - commission rule lookup fields
- Add normalized or case-insensitive indexes instead of repeated unbounded `ILIKE` searches.
- Define deterministic rule precedence using explicit priority, specificity, effective date, and stable tie-breaking.

The current percentage queries use `LIMIT 1` without ordering. This must be corrected before the backend becomes the calculation authority.

## Phase 4: Backfill And Reconciliation

Create an idempotent migration/import command that:

1. Reads all known `calc_result_storage` keys.
2. Parses each known object shape.
3. Maps companies, agents, regions, users, and codes.
4. Creates canonical policy and workflow records.
5. Stores original JSON and source keys for rollback.
6. Records unmapped references and rejected rows.
7. Produces reconciliation reports.

Backfill order:

1. companies, regions, users, and agents
2. agent-company codes
3. monthly periods
4. configuration snapshots
5. imported and office-sale policies
6. commission components
7. expenses and cash books
8. renewals, tasks, bookmarks, reminders, audit, and Amex records

Run the backfill twice and verify that the second run creates no additional canonical records.

### Reconciliation reports

For every month compare:

- policy row counts
- OSAGO and voluntary counts
- company totals
- agent totals
- office commission
- agent commission
- manager/operator commission
- profit
- paid amounts
- cash closing balances
- locked status
- configuration snapshot checksum
- unmapped agents or companies
- duplicate candidates

## Phase 5: Backend API Migration

Add authenticated calculator APIs for:

- policy imports and validation
- monthly policy lists
- office sales
- monthly periods and locks
- rate sets and rules
- cash books and payments
- expenses
- renewals
- tasks and comments
- bookmarks and reminders
- audit records
- Amex top-ups

All routes must enforce user/role permissions. The existing unrestricted `/api/calc-result-storage` route must not remain the primary business-data API.

Use transactions for:

- policy save plus commission calculation
- policy payment plus cash-book update
- cash close/reopen
- month lock/unlock
- import batch creation

Use server-generated IDs and idempotency keys.

Consolidate the existing logic in `src/routes/calculate.js` with the calculator calculation service. There should be one rule engine, not separate client and backend interpretations of the same rates.

## Phase 6: Frontend Migration

Update `calcStorage` usage incrementally:

1. Replace monthly policy and office-sale reads/writes with typed API clients.
2. Keep current spreadsheet parsing and validation initially.
3. Submit normalized rows as import batches.
4. Read effective rates, policies, snapshots, locks, payments, and cash through APIs.
5. Move commission calculation authority to the backend after parity tests pass.
6. Add loading, retry, authorization, conflict, locked-month, and partial-import states.
7. Remove month-key scanning and embedded monthly arrays from the frontend.
8. Keep Excel exports and report formatting stable.

Browser-only layout state such as drag position, renewal column visibility, and sort preferences can remain in `localStorage`.

## Phase 7: Rollout Strategy

### Phase A: Schema and read-only verification

Deploy the schema, models, indexes, backfill tooling, and reconciliation reports. Do not change calculator writes yet.

### Phase B: Dual write

Write new canonical relational records while retaining JSON compatibility writes. Track failed secondary writes in an outbox or retry table.

### Phase C: Dual read and comparison

For selected months, read both sources and compare:

- policies
- sales
- commission totals
- search results
- payments
- cash balances
- lock state
- snapshot checksums

Show mismatches to administrators only.

### Phase D: Pilot cutover

Switch relational reads on for selected users and months. Monitor mismatch counts, query latency, constraint failures, and failed writes.

### Phase E: Full cutover

Switch all users to relational APIs. Freeze direct writes to calculator JSON storage and retain it as an archive.

### Phase F: Cleanup

After the defined rollback retention period, remove compatibility reads and obsolete JSON keys through a separate migration. Drop or archive duplicate legacy tables only after business approval.

## Verification Requirements

Add tests for:

- Armenia BM groups and short-term policies
- Nairi region/BM overrides
- agent-specific overrides
- all company rates
- voluntary product rates
- exception and brand-whitelist rules
- manager and operator tiers
- MREO behavior
- rule precedence when multiple rules match
- malformed dates and numeric values
- missing agents and companies
- duplicate imports and repeated backfill
- month locking and historical snapshot behavior
- policy payment and cash close/reopen
- role and employee permissions
- transaction rollback
- search and report totals
- Excel exports

Before destructive cleanup:

1. Take a PostgreSQL backup.
2. Preserve the original JSON archive.
3. Produce a per-month reconciliation report.
4. Confirm that all accepted mismatches have business approval.

## Relevant Existing Files

- [ui/src/pages/Calculator/calculator.jsx](ui/src/pages/Calculator/calculator.jsx) - calculator state, parsers, calculations, snapshots, locks, cash, renewals, and persistence keys.
- [ui/src/pages/Calculator/calcStorage.js](ui/src/pages/Calculator/calcStorage.js) - temporary JSONB compatibility adapter.
- [src/models/InsurancePolicy.js](src/models/InsurancePolicy.js) - legacy policy model.
- [src/models/Agent.js](src/models/Agent.js) - current agent and insurer-code model.
- [src/models/Company.js](src/models/Company.js) - company defaults.
- [src/models/AgentPercentage.js](src/models/AgentPercentage.js) - agent exception rates.
- [src/models/CompanyPercentage.js](src/models/CompanyPercentage.js) - company exception rates.
- [src/models/Region.js](src/models/Region.js) - region reference data.
- [src/models/CalcResultStorage.js](src/models/CalcResultStorage.js) - JSONB key/value persistence.
- [src/routes/calc_result_storage.js](src/routes/calc_result_storage.js) - current calculator storage API.
- [src/routes/calculate.js](src/routes/calculate.js) - existing CSV import and calculation logic.
- [src/services/InsurancePolicyService.js](src/services/InsurancePolicyService.js) - existing policy service.
- [src/database.js](src/database.js) - database initialization and implicit schema alteration.
- [migrations](migrations) - explicit migration location.
- [server.js](server.js) and [src/middleware/auth.js](src/middleware/auth.js) - API mounting and authorization.

## Further Considerations

- Add a workspace or tenant boundary now, even if only one company currently uses the application.
- Preserve `source` and `legacy_id` when merging old policy records. Do not destructively deduplicate without business review.
- Rotate the database credential if the plaintext password in `config/config.json` has been exposed.
- Update stale project metadata describing the application as SQLite after PostgreSQL migration work is complete.
