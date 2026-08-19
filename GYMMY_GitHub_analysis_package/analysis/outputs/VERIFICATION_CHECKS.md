# Verification checks for the final public analysis package

The final terminology update does not change the study values. Key anchors are:

## Acceptance: control vs. failure
- Control: M = 5.31, SD = 0.95
- Failure: M = 4.97, SD = 0.99
- Mean change = -0.34
- Paired test: t(39) = -2.94, p = .006, Cohen's dz = .46

## Failure Recognition
- Hardware: 14/20 = 70%
- Interaction: 5/20 = 25%
- Early: 7/20 = 35%
- Late: 12/20 = 60%
- Total recognized: 19/40 = 47.5%

## Two-way ANOVA on Acceptance change
- Failure type: F(1,36) = 0.326, p = .572, partial eta squared = .009
- Failure timing: F(1,36) = 0.314, p = .579, partial eta squared = .009
- Interaction: F(1,36) = 1.909, p = .176, partial eta squared = .050

Run `python verify_public_outputs.py` to check these values directly from the public CSV/output files.
