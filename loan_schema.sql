-- ============================================================
--  LOAN MONITORING SYSTEM - Database Schema & Sample Data
-- ============================================================

CREATE TABLE IF NOT EXISTS branches (
    branch_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    branch_name   TEXT NOT NULL,
    loan_officer  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS customers (
    customer_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name     TEXT NOT NULL,
    phone_number  TEXT NOT NULL,
    email         TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS loans (
    loan_id           TEXT PRIMARY KEY,
    customer_id       INTEGER NOT NULL,
    branch_id         INTEGER NOT NULL,
    amount_borrowed   REAL NOT NULL,
    outstanding_balance REAL NOT NULL,
    due_date          DATE NOT NULL,
    loan_status       TEXT NOT NULL CHECK(loan_status IN ('Active', 'Overdue', 'Paid')),
    FOREIGN KEY (customer_id) REFERENCES customers(customer_id),
    FOREIGN KEY (branch_id)   REFERENCES branches(branch_id)
);

-- ============================================================
--  SAMPLE DATA
-- ============================================================

INSERT INTO branches (branch_name, loan_officer) VALUES
    ('Nairobi CBD Branch',    'James Mwangi'),
    ('Westlands Branch',      'Sarah Otieno'),
    ('Mombasa Road Branch',   'Peter Kamau'),
    ('Upperhill Branch',      'Grace Wanjiku');

INSERT INTO customers (full_name, phone_number, email) VALUES
    ('Alice Njoroge',    '+254 712 345 678', 'alice.njoroge@email.com'),
    ('Brian Ochieng',    '+254 723 456 789', 'brian.ochieng@email.com'),
    ('Catherine Mutua',  '+254 734 567 890', 'catherine.mutua@email.com'),
    ('Daniel Kariuki',   '+254 745 678 901', 'daniel.kariuki@email.com'),
    ('Esther Wambui',    '+254 756 789 012', 'esther.wambui@email.com'),
    ('Francis Ndegwa',   '+254 767 890 123', 'francis.ndegwa@email.com'),
    ('Grace Akinyi',     '+254 778 901 234', 'grace.akinyi@email.com'),
    ('Henry Kimani',     '+254 789 012 345', 'henry.kimani@email.com'),
    ('Irene Chebet',     '+254 790 123 456', 'irene.chebet@email.com'),
    ('John Muthoni',     '+254 701 234 567', 'john.muthoni@email.com');

-- Loans with due dates spread around today (for demo purposes, dates are relative)
INSERT INTO loans (loan_id, customer_id, branch_id, amount_borrowed, outstanding_balance, due_date, loan_status) VALUES
    ('LN-2024-0101', 1, 1, 150000.00, 45000.00,  date('now', '+2 days'),  'Active'),
    ('LN-2024-0102', 2, 2, 200000.00, 200000.00, date('now', '-1 days'),  'Overdue'),
    ('LN-2024-0103', 3, 3, 85000.00,  30000.00,  date('now', '+5 days'),  'Active'),
    ('LN-2024-0104', 4, 1, 320000.00, 100000.00, date('now', '+7 days'),  'Active'),
    ('LN-2024-0105', 5, 4, 50000.00,  50000.00,  date('now', '-3 days'),  'Overdue'),
    ('LN-2024-0106', 6, 2, 175000.00, 60000.00,  date('now', '+1 days'),  'Active'),
    ('LN-2024-0107', 7, 3, 420000.00, 420000.00, date('now', '+3 days'),  'Active'),
    ('LN-2024-0108', 8, 4, 95000.00,  25000.00,  date('now', '+6 days'),  'Active'),
    ('LN-2024-0109', 9, 1, 260000.00, 80000.00,  date('now', '+30 days'), 'Active'),  -- Not due in 7 days
    ('LN-2024-0110', 10,2, 110000.00, 110000.00, date('now', '+14 days'), 'Active'); -- Not due in 7 days
