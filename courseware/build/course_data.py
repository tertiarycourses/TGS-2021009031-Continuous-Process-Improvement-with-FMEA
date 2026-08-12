"""Single source of truth for TGS-2021009031 courseware."""

TITLE = "Continuous Process Improvement with FMEA"
SHORT_TITLE = "Continuous Process Improvement with FMEA"
COURSE_CODE = "TGS-2021009031"
VERSION = "v11.0"
VERSION_DATE = "13 August 2026"
ORG = "Tertiary Infotech Academy Pte Ltd"
UEN = "UEN: 201200696W"
TRAINER = "Dr. Alfred Ang"
DAYS = 2
COURSE_URL = "https://www.tertiarycourses.com.sg/wsq-continuous-process-improvement-with-fmea.html"
GITHUB_URL = "https://github.com/tertiarycourses/TGS-2021009031-Continuous-Process-Improvement-with-FMEA"
LMS_URL = "https://lms-tms.tertiaryinfotech.com/"

TSC_TITLE = "Continuous Process Improvement"
TSC_CODE = "ELE-PIN-4001-1.1"

KNOWLEDGE = [
    "K1: Identification of continuous improvement opportunities",
    "K2: Direct continuous improvement activities",
    "K3: Management of systems and processes to facilitate continuous improvement",
    "K4: Monitoring and evaluation of continuous improvement processes",
]

ABILITIES = [
    "A1: Manage and implement continuous improvement systems",
    "A2: Promote team support and coach team members toward continuous improvement",
    "A3: Gather, access and record production information to track continuous improvements",
    "A4: Analyse, synthesise and interpret information",
    "A5: Design and apply improvement tools and strategies",
    "A6: Monitor and evaluate improvement processes",
    "A7: Identify continuous improvement opportunities",
    "A8: Discuss and direct continuous improvement activities",
    "A9: Manage systems and processes to facilitate continued improvement",
    "A10: Evaluate improvement processes",
]

LEARNING_OUTCOMES = [
    "ELO1: Create a continuous process improvement system.",
    "ELO2: Perform root cause analysis for continuous process improvement.",
    "ELO3: Design and apply Failure Mode and Effects Analysis (FMEA).",
    "ELO4: Identify and direct improvement opportunities.",
    "ELO5: Manage and evaluate improvement with a Process Control Plan.",
]

CASE_STUDY = {
    "company": "Orion Precision Electronics (OPE)",
    "process": "Surface-mount technology (SMT) assembly of industrial sensor-controller boards",
    "problem": "Rising solder defects, rework, late lots and two customer returns after a new product ramp-up",
    "goal": "Reduce first-pass defects, stabilise cycle time and prevent recurrence through a living PFMEA and Control Plan",
}

TOPICS = [
    dict(num=1, code="01", title="Introduction to Continuous Process Improvement",
         subtitle="Culture · governance · PDCA · process definition",
         weighting="ELO1 · A1 · A2 · K1",
         concepts=[
             ("System, not event", "Continuous improvement is a managed cycle of evidence, action, learning and standardisation."),
             ("People closest to work", "Operators and technicians surface weak signals and practical constraints."),
             ("PDCA discipline", "Test a defined change, compare evidence and standardise only what works."),
             ("Shared process view", "SIPOC and a process map establish boundaries before analysis begins."),
         ]),
    dict(num=2, code="02", title="Root Cause Analysis",
         subtitle="Evidence · Pareto · Fishbone · 5 Whys · Fault Tree",
         weighting="ELO2 · A3 · A4",
         concepts=[
             ("Describe the gap", "A good problem statement names the measure, location, timing and magnitude without guessing the cause."),
             ("Separate cause from symptom", "The customer effect, process failure and causal mechanism belong at different levels."),
             ("Triangulate", "Use data prioritisation, structured brainstorming and verification together."),
             ("Mechanism over blame", "Useful causes describe how the process creates failure and point toward controllable action."),
         ]),
    dict(num=3, code="03", title="Design and Apply FMEA",
         subtitle="Scope · function · failure chain · risk · action · learning",
         weighting="ELO3 · A5 · A6 · K3",
         concepts=[
             ("Prospective thinking", "FMEA asks what could fail before the next customer or process loss occurs."),
             ("Failure chain", "Function → failure mode → effect → cause → controls makes risk traceable."),
             ("Risk-informed action", "Severity, occurrence and detection guide judgement; RPN alone never grants permission to ignore risk."),
             ("Living knowledge", "Actions, verification, re-rating and review triggers keep the FMEA current."),
         ]),
    dict(num=4, code="04", title="Identify and Direct Improvement Opportunities",
         subtitle="Prioritise · select · pilot · coach · implement",
         weighting="ELO4 · A7 · A8 · K2",
         concepts=[
             ("Target the failure mechanism", "Prevention is stronger than inspection when the cause can be removed or constrained."),
             ("Choose transparently", "Decision criteria make trade-offs visible and reduce preference-driven selection."),
             ("Pilot before scale", "A bounded test checks benefit, side effects, capability and adoption."),
             ("Direct with ownership", "Every action needs an owner, due date, success measure and escalation path."),
         ]),
    dict(num=5, code="05", title="Manage Improvement with a Process Control Plan",
         subtitle="Variation · capability · SPC · reaction · review",
         weighting="ELO5 · A9 · A10 · K4",
         concepts=[
             ("Control the critical few", "Translate high-risk PFMEA causes and characteristics into practical controls."),
             ("Know the variation", "Common-cause variation needs system improvement; special causes need investigation."),
             ("React by design", "A Control Plan must state who acts, what is contained and how the process restarts."),
             ("Close the loop", "Control results, incidents and changes feed back into the PFMEA and standard work."),
         ]),
]

# Slide specs: tiles, cards, flow, two, big, pdca, pareto, fishbone, faulttree,
# failurechain, riskmatrix, controlchart, controlplan. Detailed procedures remain in the LG.
CONTENT_SLIDES = {
1: [
    ("big", "Improvement is a management system", "A workshop creates ideas; a system creates repeatable learning, ownership and control."),
    ("tiles", "What continuous improvement changes", [
        ("Performance", "Reduce defects, delay, cost and instability against a defined baseline."),
        ("Capability", "Build reliable processes that repeatedly meet customer and regulatory needs."),
        ("Learning", "Turn experiments, failures and operator knowledge into standard work."),
        ("Culture", "Make problem visibility safe and improvement part of normal work."),
    ]),
    ("two", "Continuous improvement and breakthrough change", ["Frequent, evidence-led adjustments", "Lower implementation risk", "Builds habits and process knowledge"], ["Larger redesign or technology shift", "Higher uncertainty and coordination", "May be necessary when the current system is structurally incapable"]),
    ("cards", "Lean, Six Sigma and FMEA", [
        ("Lean", ["Improve flow", "Remove non-value work", "Shorten lead time"]),
        ("Six Sigma", ["Reduce variation", "Use data to test causes", "Improve capability"]),
        ("FMEA", ["Anticipate failure", "Prioritise prevention", "Link risk to controls"]),
    ]),
    ("tiles", "A culture that can improve", [
        ("Will", "Leaders make priorities and trade-offs explicit."),
        ("Skill", "Teams know how to frame, analyse, test and control."),
        ("Authority", "Owners can act within agreed boundaries."),
        ("Guidance", "Coaching and standards turn tools into good judgement."),
    ]),
    ("flow", "Continuous improvement governance", ["Sponsor sets direction", "Process owner defines gap", "Cross-functional team analyses", "Action owners implement", "Governance reviews evidence"]),
    ("pdca", "PDCA turns change into learning", ["Plan: define gap, target and hypothesis", "Do: test at controlled scale", "Check: compare evidence and side effects", "Act: standardise, adjust or abandon"]),
    ("tiles", "Plan with evidence", [
        ("Problem", "Current versus required performance, with scope and time."),
        ("Baseline", "Trusted operational data and agreed definitions."),
        ("Target", "Specific outcome, measure and time horizon."),
        ("Hypothesis", "Why the proposed countermeasure should change the result."),
    ]),
    ("tiles", "Team roles in improvement", [
        ("Sponsor", "Removes barriers and protects the business priority."),
        ("Process owner", "Owns performance and the post-project control system."),
        ("Facilitator", "Keeps the method rigorous and participation balanced."),
        ("Operators and experts", "Provide actual-work knowledge and test feasibility."),
        ("Data owner", "Protects definitions, lineage and evidence quality."),
        ("Action owner", "Implements, verifies and documents a countermeasure."),
    ]),
    ("flow", "SIPOC establishes the analysis boundary", ["Suppliers", "Inputs", "Process", "Outputs", "Customers"]),
    ("two", "Voice of customer and voice of process", ["Needs, complaints and field effects", "Critical-to-quality expectations", "What failure means to the user"], ["Yield, cycle time and defect mechanisms", "Variation, controls and capability", "What the operation can repeatedly deliver"]),
    ("tiles", "Coaching for team support", [
        ("Ask before telling", "Use questions to surface assumptions and evidence."),
        ("Make work visible", "Use boards, owners and dates so risk is shared."),
        ("Reinforce learning", "Review what changed and why, not only the result."),
        ("Protect candour", "Treat an exposed weakness as useful information."),
    ]),
],
2: [
    ("big", "A root cause is a verified mechanism", "It is not the first plausible explanation, a correlation, or the person nearest the failure."),
    ("tiles", "Build a neutral problem statement", [
        ("What", "The measurable defect, delay or performance gap."),
        ("Where", "Process step, line, product family or customer segment."),
        ("When", "Start, pattern, shift, period or change point."),
        ("How much", "Magnitude, rate, trend and business/customer impact."),
    ]),
    ("two", "Cause, failure and consequence", ["Cause: stencil aperture partly blocked", "Mechanism: insufficient paste transfer", "Controllable through cleaning and verification"], ["Failure mode: insufficient solder joint", "Effect: intermittent signal or field return", "Observed symptom is not automatically the root cause"]),
    ("pareto", "Pareto focuses the investigation", [42, 26, 15, 9, 5, 3], ["Solder bridge", "Insufficient solder", "Misalignment", "Wrong component", "Tombstone", "Other"]),
    ("tiles", "Read a Pareto chart correctly", [
        ("Rank", "Bars order categories by frequency or impact."),
        ("Cumulative line", "Shows how quickly the categories account for the total."),
        ("Stratify", "Check shift, product and machine before combining data."),
        ("Do not overclaim", "The largest category is a priority, not yet a cause."),
    ]),
    ("fishbone", "Fishbone structures possible causes", ["People", "Machine", "Method", "Material", "Measurement", "Environment"]),
    ("tiles", "From brainstorm to evidence", [
        ("Candidate cause", "Write a specific mechanism, not a vague label."),
        ("Prediction", "State what data pattern would be expected if it were true."),
        ("Test", "Compare lots, shifts, machines, materials or controlled trials."),
        ("Decision", "Confirm, reject or keep open with the next evidence needed."),
    ]),
    ("flow", "Five Whys follows one causal chain", ["Observed failure", "Immediate mechanism", "Process condition", "System weakness", "Controllable root cause"]),
    ("two", "Strong and weak 5 Whys", ["Each answer is evidence-supported", "The chain stays on one mechanism", "The endpoint suggests a system change"], ["Jumps between unrelated branches", "Ends with 'operator error'", "Treats five as a magic number"]),
    ("faulttree", "Fault Tree Analysis tests combinations", ["Top event", "OR: any branch can cause it", "AND: conditions must combine", "Basic events with evidence"]),
    ("two", "Fishbone, 5 Whys and Fault Tree", ["Fishbone widens the search", "5 Whys deepens one likely chain", "Use when mechanisms are still being explored"], ["Fault Tree formalises logical combinations", "Supports safety and reliability reasoning", "Use when several paths can create the top event"]),
    ("tiles", "Data quality checks before conclusion", [
        ("Definition", "Are defect categories and denominators consistent?"),
        ("Measurement", "Can inspectors or sensors reproduce the classification?"),
        ("Completeness", "Are rework, scrap and escaped defects all represented?"),
        ("Context", "Are product mix, shift and process changes recorded?"),
    ]),
],
3: [
    ("big", "FMEA is prevention made visible", "It documents what might fail, why it matters, what controls exist and who will reduce the risk."),
    ("tiles", "When FMEA has the most leverage", [
        ("New design or process", "Prevent known patterns before approval or launch."),
        ("Changed use", "Reassess risk when a product or process enters a new context."),
        ("Improvement goal", "Anticipate failure introduced by the proposed change."),
        ("Before Control Plan", "Convert prioritised risks into prevention and detection controls."),
    ]),
    ("cards", "Three common FMEA views", [
        ("System FMEA", ["Interfaces and interactions", "System functions", "Cross-boundary effects"]),
        ("Design FMEA", ["Product functions", "Materials and geometry", "Design prevention"]),
        ("Process FMEA", ["How production introduces defects", "6M causes", "Feeds the Control Plan"]),
    ]),
    ("two", "Prospective FMEA and retrospective RCA", ["Asks: what could fail?", "Acts before the next loss", "Uses functions, experience and change knowledge"], ["Asks: what happened and why?", "Acts after a failure signal", "Verified findings should update the FMEA"]),
    ("flow", "Seven-step FMEA method", ["Plan & prepare", "Structure", "Function", "Failure", "Risk", "Optimise", "Document results"]),
    ("tiles", "Step 1 — planning and preparation", [
        ("Purpose", "Name the decision the FMEA must support."),
        ("Scope", "Set product/process boundaries and interfaces."),
        ("Team", "Include design, process, quality, maintenance and operators."),
        ("Inputs", "Gather flow maps, complaints, controls, changes and lessons."),
    ]),
    ("tiles", "Steps 2–3 — structure and function", [
        ("Structure", "Break the process into logical steps and focus elements."),
        ("Interfaces", "Show upstream, downstream and support interactions."),
        ("Function", "Use verb–noun language with a measurable requirement."),
        ("Requirement", "State the expected output, tolerance or customer need."),
    ]),
    ("failurechain", "Step 4 — build the failure chain", ["Function", "Failure mode", "Effect", "Cause", "Current controls"]),
    ("tiles", "Write useful failure statements", [
        ("Failure mode", "Loss, partial, intermittent, excessive or unintended function."),
        ("Effect", "What the next operation, system or customer experiences."),
        ("Cause", "A specific mechanism or process condition that creates the mode."),
        ("Control", "What prevents the cause or detects cause/mode before escape."),
    ]),
    ("two", "Prevention and detection controls", ["Remove or reduce the cause", "Standard work, robust design, interlock", "Influences occurrence"], ["Reveal cause or failure before escape", "Sensor, test, inspection, alarm", "Influences detection ranking"]),
    ("riskmatrix", "Step 5 — S, O and D describe different risk dimensions", ["Severity: consequence", "Occurrence: likelihood", "Detection: ability to catch before impact"]),
    ("tiles", "Severity rating discipline", [
        ("Customer effect", "Rate the most serious credible effect for the failure mode."),
        ("Consistent scale", "Use one agreed 1–10 table and record the rationale."),
        ("Safety/regulatory", "Escalate high-severity concerns regardless of a modest RPN."),
        ("Design change", "Detection alone does not reduce the consequence severity."),
    ]),
    ("tiles", "Occurrence rating discipline", [
        ("Cause-specific", "Rate the likelihood of the stated causal mechanism."),
        ("Use evidence", "Prefer capability, defect, maintenance or comparable-history data."),
        ("Prevention matters", "Strong prevention can reduce occurrence."),
        ("Assumptions visible", "Record uncertainty and the basis for the score."),
    ]),
    ("tiles", "Detection rating discipline", [
        ("Higher means weaker", "A high D score means the failure is hard to detect."),
        ("Control-specific", "Rate the actual detection control, not intended inspection."),
        ("Before effect", "Ask whether the cause or mode is caught before customer impact."),
        ("Evidence", "Use validation, gauge reliability and escape history."),
    ]),
    ("cards", "RPN is useful—but incomplete", [
        ("Calculate", ["RPN = S × O × D", "Range 1–1000", "Useful for sorting and trending"]),
        ("Interpret", ["Ordinal ratings", "Same RPN can hide different profiles", "Record scoring rationale"]),
        ("Guardrail", ["No fixed cutoff as permission", "Severity first", "Use team judgement and Action Priority"]),
    ]),
    ("two", "Same number, different risk", ["S=10, O=2, D=2", "RPN = 40", "Potentially safety-critical effect"], ["S=5, O=4, D=2", "RPN = 40", "Different consequence and action logic"]),
    ("flow", "Step 6 — optimise in the right order", ["Eliminate or reduce the effect", "Prevent the cause / reduce occurrence", "Improve detection", "Assign owner and due date", "Verify then re-rate"]),
    ("tiles", "Strong recommended actions", [
        ("Specific", "Names the change, location and failure mechanism."),
        ("Owned", "One accountable owner and a realistic due date."),
        ("Verifiable", "Defines evidence that proves implementation and effectiveness."),
        ("Traceable", "Links to the PFMEA row, change record and Control Plan."),
    ]),
    ("flow", "Step 7 — document and sustain", ["Record decisions", "Confirm actions", "Re-score verified state", "Update Control Plan", "Set review triggers"]),
    ("tiles", "Review triggers for a living FMEA", [
        ("Engineering change", "New material, equipment, tooling, software or parameter."),
        ("Process signal", "Capability loss, special cause, defect trend or downtime."),
        ("Customer signal", "Complaint, return, field failure or changed requirement."),
        ("Learning", "RCA, audit, maintenance finding or supplier feedback."),
    ]),
    ("tiles", "FMEA pitfalls", [
        ("Scope creep", "A boundary too broad prevents meaningful depth."),
        ("Template completion", "Scoring without insight or action creates false assurance."),
        ("RPN gaming", "Adjusting ratings to cross a threshold hides risk."),
        ("Missing operators", "The team loses actual-work failure knowledge."),
        ("No ownership", "Actions without owner/date remain recommendations."),
        ("Stale document", "A frozen FMEA stops reflecting the real process."),
    ]),
],
4: [
    ("big", "The best action changes the failure mechanism", "Inspection may catch a defect; prevention changes why it occurs."),
    ("flow", "From risk to improvement portfolio", ["High-priority risk", "Countermeasure options", "Decision criteria", "Pilot and evidence", "Approved implementation"]),
    ("cards", "Countermeasure hierarchy", [
        ("Prevent", ["Eliminate the cause", "Robust design", "Poka-yoke or interlock"]),
        ("Reduce", ["Improve capability", "Control variation", "Standard work and maintenance"]),
        ("Detect", ["Automated checks", "Validated inspection", "Fast containment signal"]),
    ]),
    ("tiles", "Opportunity sources", [
        ("PFMEA", "High severity, Action Priority, occurrence and weak detection."),
        ("Pareto and RCA", "Recurring defects and verified mechanisms."),
        ("Capability and SPC", "Unstable or incapable critical characteristics."),
        ("People and customers", "Operator constraints, complaints and near misses."),
    ]),
    ("tiles", "DMAIC provides a project roadmap", [
        ("Define", "Clarify the customer, problem, scope, target and ownership."),
        ("Measure", "Validate definitions and establish a trustworthy baseline."),
        ("Analyse", "Prioritise patterns and verify causal mechanisms."),
        ("Improve", "Select, pilot and implement countermeasures."),
        ("Control", "Monitor, react, standardise and transfer ownership."),
        ("Relationship to PDCA", "Both are evidence-led cycles; DMAIC gives a detailed improvement-project structure."),
    ]),
    ("tiles", "5S supports stable process conditions", [
        ("Sort", "Remove items that are not needed for the work."),
        ("Set in order", "Make the correct item and sequence obvious."),
        ("Shine", "Clean while inspecting for early abnormality."),
        ("Standardise", "Define a repeatable condition and visual method."),
        ("Sustain", "Audit, coach and correct drift as normal work."),
    ]),
    ("two", "DOE tests factors and interactions", ["Plan combinations of factor levels", "Estimate main effects and interactions", "Use randomisation, replication and blocking where appropriate"], ["Do not change many factors informally", "Confirm measurement and process stability first", "Use statistical expertise for safety-critical or complex designs"]),
    ("tiles", "Evaluation criteria", [
        ("Risk reduction", "Expected change to severity, occurrence or detection."),
        ("Effectiveness", "Strength of evidence and fit to the causal mechanism."),
        ("Feasibility", "Skills, time, technology and operational constraints."),
        ("Cost and speed", "Investment, lifecycle cost and time to benefit."),
        ("Side effects", "New failure modes, workload or compliance concerns."),
        ("Sustainability", "Ability to standardise, train, monitor and maintain."),
    ]),
    ("two", "Weighted criteria matrix and Pugh matrix", ["Scores options against weighted criteria", "Useful when priorities differ in importance", "Shows a total and sensitivity to weights"], ["Compares each option with a datum", "Uses better / same / worse reasoning", "Useful for concept screening and discussion"]),
    ("tiles", "Cost-benefit without false precision", [
        ("Benefit", "Avoided scrap, rework, downtime, complaints and risk exposure."),
        ("Cost", "Capital, validation, training, maintenance and transition loss."),
        ("Time", "Separate one-off savings from recurring annual impact."),
        ("Uncertainty", "Show assumptions and ranges instead of one confident number."),
    ]),
    ("flow", "Pilot learning cycle", ["Hypothesis and success measures", "Safe scope and baseline", "Run with controls", "Compare result and side effects", "Decide scale / revise / stop"]),
    ("tiles", "A credible pilot", [
        ("Representative", "Covers the relevant product, shift and failure conditions."),
        ("Contained", "Limits customer and production risk during the test."),
        ("Measured", "Uses defined input, process and outcome measures."),
        ("Reversible", "Includes a stop rule and recovery plan."),
    ]),
    ("cards", "Poka-yoke strategies", [
        ("Contact", ["Physical fit or sensor", "Detects shape/position", "Prevents wrong assembly"]),
        ("Fixed value", ["Counts required items or actions", "Signals omissions", "Useful for kits and fasteners"]),
        ("Motion sequence", ["Enforces task order", "Interlock or checklist", "Prevents skipped steps"]),
    ]),
    ("tiles", "Direct and coach the team", [
        ("Purpose", "Explain the customer/risk reason for the change."),
        ("Practice", "Demonstrate, observe and let the operator perform."),
        ("Feedback", "Correct the method and invite actual-work concerns."),
        ("Follow-up", "Check adoption, outcome and unintended effects."),
    ]),
    ("tiles", "Implementation control", [
        ("Owner and date", "One accountable person and milestone."),
        ("Change control", "Approved specification, work instruction and training."),
        ("Verification", "Evidence the action was installed as intended."),
        ("Effectiveness", "Evidence it reduced risk without creating new failure."),
    ]),
],
5: [
    ("big", "Control is a feedback system", "A stable process detects meaningful change, contains risk and learns before the customer does."),
    ("two", "Common-cause and special-cause variation", ["Built into the current system", "Many small sources", "Improve the process as a system"], ["Specific, unusual signal", "Assignable event or condition", "Investigate and remove the cause"]),
    ("tiles", "Stability and capability answer different questions", [
        ("Stability", "Is the process behaviour predictable over time?"),
        ("Capability", "Can a stable process meet specification consistently?"),
        ("Sequence", "Check stability before interpreting capability indices."),
        ("Customer", "Specifications come from requirements; control limits come from process data."),
    ]),
    ("cards", "Cp and Cpk", [
        ("Cp", ["Potential capability", "Specification width ÷ 6σ", "Assumes centring"]),
        ("Cpk", ["Actual centring-aware capability", "Minimum distance to limit ÷ 3σ", "Cpk ≤ Cp"]),
        ("Interpret", ["Use stable data", "Confirm distribution and measurement", "Report assumptions"]),
    ]),
    ("controlchart", "A control chart separates signal from noise", ["Centre line", "Upper and lower control limits", "Time-ordered observations", "Rules for non-random patterns"]),
    ("two", "Control limits and specification limits", ["Calculated from process behaviour", "Support stability decisions", "Not changed to make the chart look better"], ["Defined by customer/design requirement", "Support conformance decisions", "Do not describe natural process variation"]),
    ("tiles", "Control-chart signals", [
        ("Point beyond a limit", "Immediate special-cause investigation."),
        ("Run on one side", "Possible process shift."),
        ("Trend", "Gradual change, wear or drift."),
        ("Cycle or clustering", "Time pattern, batching or stratification."),
    ]),
    ("controlplan", "The Control Plan operationalises the PFMEA", ["Process step / characteristic", "Specification and method", "Sample size and frequency", "Owner and record", "Reaction plan"]),
    ("tiles", "Control selection principles", [
        ("Risk-based", "Strength and frequency reflect severity, occurrence and escape risk."),
        ("At source", "Control as close as possible to the causal process condition."),
        ("Capable measurement", "The method can reliably distinguish acceptable from unacceptable."),
        ("Actionable", "A signal must trigger a defined response, not only a record."),
    ]),
    ("flow", "Reaction plan", ["Stop or contain", "Identify affected product", "Notify owner", "Correct and verify", "Release and update records"]),
    ("two", "Monitoring and evaluation", ["Leading: parameter adherence, preventive maintenance, alarm rate", "Detect drift before outcome loss", "Useful for early intervention"], ["Lagging: defect, scrap, rework, return, downtime", "Confirms realised outcome", "Useful for accountability and learning"]),
    ("tiles", "Evaluate an improvement", [
        ("Effectiveness", "Did the targeted measure improve against baseline?"),
        ("Sustainment", "Did the result persist across time, shift and product mix?"),
        ("Side effects", "Did another defect, delay or workload appear?"),
        ("Adoption", "Is standard work followed and understood?"),
        ("Risk", "Are PFMEA ratings and controls updated with evidence?"),
        ("Learning", "What should be standardised, transferred or investigated next?"),
    ]),
    ("flow", "Closed-loop improvement", ["Monitor control data", "Detect signal", "Contain and investigate", "Update PFMEA and action", "Revise Control Plan and standard work"]),
    ("tiles", "Governance review questions", [
        ("Right risk?", "Are high-severity and high-priority items visible?"),
        ("Right evidence?", "Are ratings, capability and results based on trusted data?"),
        ("Right action?", "Does the countermeasure address the verified mechanism?"),
        ("Right control?", "Will drift be detected and contained in time?"),
    ]),
],
}

ASSESSMENT = dict(
    written="Written Assessment (WA-SAQ) — 4 open-ended questions, 60 minutes, open book.",
    practical="Practical Performance (PP) — 5 connected tasks, 90 minutes, open book.",
    note="Assessment design follows the approved plan and existing papers: K1–K4 and A1–A10.",
)

SOURCES = [
    ("ASQ — Failure Mode and Effects Analysis", "https://asq.org/quality-resources/fmea"),
    ("Overview of FMEA: A Patient Safety Tool", "https://pmc.ncbi.nlm.nih.gov/articles/PMC10229026/"),
    ("Jama Software — FMEA Explained", "https://www.jamasoftware.com/requirements-management-guide/meeting-regulatory-compliance-and-industry-standards/fmea/"),
    ("Quality-One — FMEA", "https://quality-one.com/fmea/"),
    ("ScienceDirect Topics — FMEA", "https://www.sciencedirect.com/topics/engineering/failure-modes-and-effects-analysis"),
    ("IPH Hannover — FMEA", "https://www.iph-hannover.de/de/dienstleistungen/fertigungsverfahren/fmea/"),
]
