# Salary & Expense Payment Automation System (Excel VBA)

## Files
- `SalaryExpenseAutomation.bas` - Main reusable VBA module.
- `UserFormsCode.txt` - UserForm code-behind templates for all required forms.

## Setup
1. Open Excel workbook (`.xlsm`).
2. Import `SalaryExpenseAutomation.bas` into VBA project.
3. Create UserForms with names:
   - `frmEmployeeMaster`
   - `frmSalaryEntry`
   - `frmExpenseEntry`
   - `frmEmployeeSearch`
4. Add controls exactly as documented in `UserFormsCode.txt` comments.
5. Paste respective form event code into each form.
6. Run `SetupSystem` macro once.

## Security Notes
- Update `MASTER_PASSWORD` constant to your own password.
- Keep VBA project password-protected.
- MASTER sheet is locked and edited only via form-driven macros.

## Output
Generated bank file headers are strict and fixed per requirement.
