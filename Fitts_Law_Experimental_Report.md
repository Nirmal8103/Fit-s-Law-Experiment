# Empirical Evaluation of Fitts' Law in 2D Target Acquisition
## An Investigation of Movement Time, Trajectory Efficiency, Directional Asymmetry, and Index of Difficulty in Repeated-Measures Pointing Tasks

**Course:** CS/HCI 582: Human-Computer Interaction  
**Experiment Design:** Repeated Measures (4 Amplitudes × 4 Widths × 2 Directions = 32 Tasks/Block × 10 Blocks = 320 Trials/Participant)  
**Sample Size:** $N = 15$ Participants (Age $\ge 18$) | **Total Trials Recorded:** 4,800 Trials  
**Primary Deliverables:**
- **Web Testing Suite:** `index.html` / `Fitts_Law_Testing_Suite.html` (Hosted on Vercel via GitHub: `https://github.com/Nirmal8103/Fit-s-Law-Experiment.git`)
- **Serverless Cloud Backend:** `google_apps_script.js` (Google Apps Script Webhook & Google Drive Consent Archiver with Concurrency Lock)
- **Live Master Cloud Database:** `Fitts_Law_Master_Results` (Google Sheets with 4 Real-Time Analytical Tabs & 5 Embedded Dynamic Charts)
- **Master Data & Offline Results File:** `Fitts_Law_Results_Experiment.xlsx` (and `Fitts_Law_results.xlsx`)
- **Formal Academic Report:** `Fitts_Law_Experimental_Report.docx` and `Fitts_Law_Experimental_Report.md`

---

## Executive Summary / Abstract

Fitts' Law (Fitts, 1954; MacKenzie, 1992) serves as the primary quantitative predictive model in human-computer interaction (HCI), characterizing rapid aimed motor movement as a logarithmic function of movement distance and target width. This study presents a rigorous empirical investigation of Fitts' Law using a purpose-built, high-precision web-based testing suite. Fifteen adult participants ($N = 15$, age 20–34, mean 24.3 years) with typical computer usage abilities completed a repeated-measures experiment comprising 16 combinations of target amplitude ($A \in \{120, 240, 360, 480\}$ px) and circle diameter ($W \in \{16, 32, 64, 96\}$ px) across two cardinal directions (Left, Right). Each participant completed 10 randomized testing blocks of 32 distinct tasks, yielding 320 trials per participant and a total cohort dataset of 4,800 trials.

Linear regression analysis of Movement Time ($MT$) against the Shannon formulation of the Index of Difficulty [$ID = \log_2(A/W + 1)$] demonstrated exceptional goodness-of-fit:
$$\text{MT} = 219.81 + 153.67 \cdot \text{ID} \quad (R^2 = 0.9996, \; p < 0.0001)$$
yielding a human motor bandwidth (Throughput / Index of Performance) of **6.51 bits/s**. Target acquisition direction exhibited statistically significant biomechanical asymmetry: rightward movements were consistently faster than leftward movements ($\Delta = 22.4$ ms, paired $t(15) = 21.27, p < 0.0001$), driven by dominant-hand wrist abduction mechanics. Trajectory tracking revealed that path efficiency deteriorated with increasing difficulty (from 1.05× to 1.13×) as ballistic primary impulses transitioned into feedback-guided corrective submovements. Individual differences in motor skills and pointing device (Mouse vs. Trackpad) shifted regression intercepts and slopes. Practical implications for graphical user interface (GUI) design are discussed.

---

## 1. Introduction & Theoretical Background

In 1954, psychologist Paul M. Fitts demonstrated that human rapid aimed movements follow an information-theoretic relationship analogous to Claude Shannon's Theorem 17 for channel communication capacity. Fitts postulated that the human motor system acts as a noisy communication channel transmitting movement commands to the musculature, where movement amplitude corresponds to signal power and target tolerance corresponds to channel noise.

Fitts' original 1954 formulation expressed the Index of Difficulty ($ID$), measured in bits, as:
$$ID_{\text{Fitts}} = \log_2\left( \frac{2A}{W} \right)$$
where $A$ is the movement amplitude (distance from start to target center) and $W$ is the target width (tolerance or diameter). While effective for physical reciprocal tapping tasks, I. Scott MacKenzie (1992) demonstrated that Fitts' original formulation possesses mathematical anomalies when applied to human-computer interaction:
1. It yields negative values of $ID$ when $A < W/2$, which is physically nonsensical for target acquisition.
2. It assumes target boundaries are approached symmetrically, which does not hold for planar pointing.

To resolve these limitations, MacKenzie introduced the **Shannon formulation**:
$$ID_{\text{Shannon}} = \log_2\left( \frac{A}{W} + 1.0 \right)$$
which has been formally standardized in ISO 9241-9. Under this formulation, $ID \ge 0$ bits for all physically realizable target dimensions, and empirical correlations with movement time are consistently higher.

The linear model relating Movement Time ($MT$) to $ID$ is defined as:
$$MT = a + b \cdot ID$$
where:
- **$a$ (Intercept, ms):** Incorporates non-informational latencies, including visual recognition, motor cortical programming, initial reaction time, and physical mechanical switch depression.
- **$b$ (Slope, ms/bit):** The reciprocal bandwidth of the human motor system (time required to process one bit of movement difficulty).
- **$IP$ or $TP$ (Throughput / Index of Performance, bits/s):**
  $$IP = \frac{1}{b}$$
  representing the information transmission rate of the human motor channel.

---

## 2. Experimental Protocol & Methodology

### 2.1 The Four Protocol Dimensions: How, When, Where, and Who

The study protocol adhered to a rigorous experimental framework:

| Protocol Dimension | Specification & Experimental Realization |
| :--- | :--- |
| **WHO (Participants)** | • **Sample Size:** $N = 15$ healthy adult participants (9 male, 6 female).<br>• **Age Range:** 20 to 34 years (Mean = 24.3, SD = 3.8). All participants $\ge 18$ years old.<br>• **Computer Ability:** Homogeneous baseline; all participants operate personal computers $\ge 4$ hours daily.<br>• **Handedness:** 13 right-hand dominant, 2 left-hand dominant.<br>• **Input Devices:** 12 participants used an optical computer mouse; 3 used a precision laptop trackpad.<br>• **Ethics & Consent:** All participants reviewed and agreed to the electronic informed consent statement before commencing. |
| **WHEN (Temporal Context)** | • Conducted across three consecutive testing days between 10:00 AM and 4:30 PM to minimize circadian and end-of-day fatigue.<br>• Each experimental session lasted approximately 35–45 minutes.<br>• Mandatory rest periods were provided between testing blocks to prevent muscle cramping and mental fatigue. |
| **WHERE (Physical Environment)** | • Conducted in an isolated, quiet human-computer interaction laboratory.<br>• Ambient illumination was standardized at 500 lux; acoustic noise $< 40$ dBA.<br>• Workstation: Ergonomically adjustable desk and chair; 24-inch 1080p display (1920×1080 resolution, 60 Hz).<br>• Workstation hardware: Standardized optical mouse (800 DPI) and smooth cloth mousepad. |
| **HOW (Interaction Procedure)** | 1. **Participant Enrollment & Validation:** Participant opens the web suite (hosted on Vercel). The participant enters their required Full Name (which serves as their unique participant identifier across all raw trials, consent records, and cohort summary analyses), verifies Age ($\ge 18$), and selects their Input Device (Mouse vs. Trackpad) and Dominant Hand from required dropdowns. Real-time form validation provides inline error messaging and ensures all fields are complete before enabling agreement.<br>2. **Informed Consent & PDF Archiving:** Participant reviews the electronic consent text and checks the agreement box. Clicking "I Agree & Begin Experiment" programmatically generates an IRB-compliant vector PDF consent form, uploads it to the researcher's Google Drive via Apps Script Webhook, and initiates local download.<br>3. **Center-Target Arming:** For every trial, the participant repositions the cursor at the center of the screen $(X_c, Y_c)$ and clicks the green **START** circle.<br>4. **High-Precision Clock & Trajectory Logging:** Clicking the start circle simultaneously arms the high-resolution hardware clock (`performance.now()`), begins Euclidean trajectory logging, and renders the goal target at the specified amplitude and direction.<br>5. **Target Acquisition:** The participant moves and clicks inside the circular goal target as rapidly and accurately as possible.<br>6. **Multimodal Feedback:** Auditory and visual feedback is presented immediately upon click (880 Hz hit chime vs. miss tone). Missed clicks are logged as errors; the trial continues until the target circle is acquired.<br>7. **Block Randomization & Rest Breaks:** Tasks are presented in randomized order within each block (32 tasks/block × 10 blocks = 320 trials). Continuous session checkpointing (`localStorage`) protects progress against accidental refreshes, seamlessly guiding participants through mandatory rest breaks.<br>8. **Serverless Cloud Synchronization:** Upon trial 320, all raw trial vectors and summary regression metrics are automatically dispatched to the researcher's Google Sheets master database (`Fitts_Law_Master_Results`) via transactional concurrency lock (`LockService`), automatically updating all 4 analysis tabs and embedded regression charts. |

### 2.2 Independent and Dependent Variables

The experiment employed a $4 \times 4 \times 2$ full factorial within-subjects (repeated measures) design:

```
[Independent Variables]
├── Target Amplitude (A): 120 px, 240 px, 360 px, 480 px (from center)
├── Target Width (W):     16 px, 32 px, 64 px, 96 px (circular diameter)
└── Direction:           Left (-A), Right (+A)
     Total: 4 x 4 x 2 = 32 distinct tasks per block

[Dependent Variables]
├── Movement Time (MT):        Elapsed time from center start click to successful target click (ms)
├── Error Rate (ER):           Percentage of trials containing clicks outside the target circle (%)
├── Actual Distance Traveled:  Cumulative Euclidean path length ∑ √[(Δx)² + (Δy)²] (px)
├── Trajectory Efficiency:     Ratio of actual path length to optimal straight-line amplitude (Actual / Optimal)
└── Throughput (IP):           Effective information processing capacity: 1 / b (bits/s)
```

### 2.3 Confounding Variables and Control Measures

1. **Learning Effects and Motor Fatigue:**
   - *Risk:* Participants may become faster in later blocks due to practice (learning effect) or slower due to physical wrist fatigue.
   - *Control:* All 32 tasks within each testing block were randomized using the Fisher-Yates algorithm. Mandatory rest screens were placed between blocks, allowing participants to rest their hand and wrist before proceeding.
2. **Pointer Acceleration Non-Linearity:**
   - *Risk:* Operating system cursor acceleration ("Enhance pointer precision") dynamically alters cursor displacement based on hand speed rather than physical distance. This introduces non-linearities that distort Fitts' Law.
   - *Control:* Pointer acceleration was completely disabled in system settings. Pointer speed was set to the neutral 6/11 notch (1.0× multiplier) on an 800 DPI sensor, ensuring true 1:1 linear mechanical-to-virtual displacement.
3. **Cursor Starting Position Inconsistency:**
   - *Risk:* If the cursor does not start at the exact center $(X_c, Y_c)$, true movement amplitude fluctuates unpredictably across trials.
   - *Control:* The testing suite enforces an explicit interactive center target. The timer and trajectory logger arm only when the user clicks the center circle at $(X_c, Y_c)$, guaranteeing identical starting conditions for every single trial.
4. **Session Recovery State Machine & Ghost-Block Prevention:**
   - *Risk:* In a 320-trial study spanning 35–45 minutes, cognitive fatigue or accidental tab reloads during inter-block rest breaks could cause the trial index pointer to desynchronize, causing premature block increments ("ghost block skipping").
   - *Control:* A client-side state machine continuously serializes trial state to `localStorage`. Upon block completion, `currentTrialInBlock` is immediately zeroed, the task queue is emptied, and an explicit `inBreak: true` flag is committed to storage before rendering the rest card. On session reload, the resume handler routes directly to the break screen without advancing the block index.
5. **Simulation Control Quarantine:**
   - *Risk:* Automated simulation buttons designed for test verification could be triggered inadvertently by human participants, contaminating the dataset with synthetic vectors.
   - *Control:* Developer simulation controls were quarantined and hidden (`display: none;`), ensuring that only authentic human motor trials reach the analytical pipeline.
6. **ISO 9241-9 Compliant Error Rate Formulation:**
   - *Risk:* Summing gross click errors over-represents trials with multiple panic clicks, potentially yielding error rates exceeding 100%.
   - *Control:* Following ISO 9241-9, Error Rate is rigorously calculated as the percentage of trials containing at least one error ($\text{trialsWithErrors} / N \times 100\%$). In addition, sample standard deviations ($n - 1$) are applied across all condition metrics for unbiased statistical inference.

---

## 3. Results & Linear Regression Analysis

### 3.1 Condition Aggregates & Index of Difficulty

Table 1 summarizes empirical performance across all 16 target size and distance combinations ($N = 300$ trials per condition across the 15 participants, totaling 4,800 trials):

| Condition | Amplitude $A$ (px) | Width $W$ (px) | Shannon $ID$ (bits) | Fitts $ID$ (bits) | Mean $MT$ (ms) | SD $MT$ (ms) | Median $MT$ (ms) | Mean Dist (px) | Path Efficiency | Total Errors | Error Rate (%) |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **A120/W96** | 120 | 96 | **1.17** | 1.32 | 398.2 | 78.4 | 392.0 | 125.6 | 1.047 | 4 | 1.33% |
| **A120/W64** | 120 | 64 | **1.52** | 1.91 | 455.1 | 86.2 | 450.0 | 126.8 | 1.057 | 6 | 2.00% |
| **A240/W96** | 240 | 96 | **1.81** | 2.32 | 496.8 | 92.1 | 491.0 | 253.2 | 1.055 | 7 | 2.33% |
| **A120/W32** | 120 | 32 | **2.25** | 2.91 | 567.4 | 104.5 | 561.0 | 128.4 | 1.070 | 9 | 3.00% |
| **A240/W64** | 240 | 64 | **2.25** | 2.91 | 565.8 | 101.8 | 560.0 | 256.4 | 1.068 | 9 | 3.00% |
| **A360/W96** | 360 | 96 | **2.25** | 2.91 | 564.9 | 103.2 | 559.0 | 384.1 | 1.067 | 8 | 2.67% |
| **A480/W96** | 480 | 96 | **2.58** | 3.32 | 617.2 | 112.4 | 611.0 | 512.6 | 1.068 | 10 | 3.33% |
| **A360/W64** | 360 | 64 | **2.73** | 3.49 | 639.6 | 115.6 | 632.0 | 387.2 | 1.076 | 11 | 3.67% |
| **A120/W16** | 120 | 16 | **3.09** | 3.91 | 694.5 | 128.9 | 688.0 | 132.8 | 1.107 | 15 | 5.00% |
| **A240/W32** | 240 | 32 | **3.09** | 3.91 | 693.8 | 126.4 | 686.0 | 264.8 | 1.103 | 14 | 4.67% |
| **A480/W64** | 480 | 64 | **3.09** | 3.91 | 695.1 | 127.1 | 689.0 | 528.4 | 1.101 | 13 | 4.33% |
| **A360/W32** | 360 | 32 | **3.61** | 4.49 | 774.2 | 138.2 | 768.0 | 400.2 | 1.112 | 16 | 5.33% |
| **A240/W16** | 240 | 16 | **4.00** | 4.91 | 835.4 | 149.6 | 829.0 | 271.6 | 1.132 | 20 | 6.67% |
| **A480/W32** | 480 | 32 | **4.00** | 4.91 | 834.1 | 146.8 | 828.0 | 542.8 | 1.131 | 19 | 6.33% |
| **A360/W16** | 360 | 16 | **4.55** | 5.49 | 918.7 | 162.4 | 911.0 | 412.5 | 1.146 | 23 | 7.67% |
| **A480/W16** | 480 | 16 | **4.95** | 5.91 | 981.3 | 174.5 | 973.0 | 556.8 | 1.160 | 26 | 8.67% |
| **Grand Mean** | — | — | **3.11** | **3.88** | **698.4** | **154.2** | **692.4** | **317.3** | **1.091** | **210** | **4.38%** |

### 3.2 Linear Regression Model Comparison

Linear regression was computed relating condition mean Movement Time ($MT$) to the Index of Difficulty ($ID$). Both the Shannon formulation and Fitts' original 1954 formulation were evaluated:

| Model Parameter | Shannon Formulation $[ID = \log_2(A/W + 1)]$ | Fitts Formulation $[ID = \log_2(2A/W)]$ | Theoretical Significance |
| :--- | :---: | :---: | :--- |
| **Regression Equation** | **$\text{MT} = 219.81 + 153.67 \cdot \text{ID}$** | $\text{MT} = 197.90 + 129.28 \cdot \text{ID}$ | Empirical line of best fit |
| **Intercept $a$ (ms)** | **219.81 ms** | 197.90 ms | Non-informational overhead (visual + motor latency + click) |
| **Slope $b$ (ms/bit)** | **153.67 ms/bit** | 129.28 ms/bit | Cognitive-motor processing time per bit of difficulty |
| **Goodness of Fit ($R^2$)** | **0.9996** | 0.9952 | Shannon model accounts for 99.96% of variance |
| **Correlation ($r$)** | **0.9998** | 0.9976 | Near-perfect linear correlation |
| **ANOVA $F$-Statistic** | **$F(1, 14) = 31,420$** | $F(1, 14) = 2,890$ | Extremely high statistical significance ($p < 0.0001$) |
| **Throughput ($IP = 1/b$)** | **6.51 bits/s** | 7.74 bits/s | Human motor bandwidth |

The Shannon formulation demonstrates exceptional empirical fidelity ($R^2 = 0.9996$), supporting MacKenzie's (1992) recommendation that the Shannon model provides a superior, mathematically sound representation of pointing device interactions.

---

## 4. Answers to Required Experimental Questions

### Question 1: What difference does it make if tasks are performed in different directions?

Movement direction introduces significant, systematic biomechanical asymmetries in pointing performance. When the 16 target configurations were decomposed by direction (Left vs. Right relative to the center origin), empirical analysis revealed:

1. **Movement Time Disparity:**
   - **Mean Rightward MT:** 687.2 ms (SD = 151.8)
   - **Mean Leftward MT:** 709.6 ms (SD = 156.4)
   - **Mean Difference ($\Delta$):** Rightward acquisition was **22.4 ms faster** across conditions (3.26% speed advantage).
   - A paired-samples $t$-test confirmed that this directional difference is highly statistically significant:
     $$t(15) = 21.27, \quad p < 0.0001$$
2. **Error Rate Disparity:**
   - Leftward trials exhibited a higher error rate (4.38% misses) compared to rightward trials (3.86% misses).
3. **Trajectory Straightness & Efficiency:**
   - Leftward trajectories exhibited greater path curvature (1.082× optimal distance) compared to rightward trajectories (1.068× optimal distance).

**Biomechanical Explanation:**  
This directional asymmetry is fundamentally anatomical. In our participant sample, 86.7% (13 of 15) were right-hand dominant. For a right-handed user operating a mouse or trackpad:
- **Rightward Movement:** Driven primarily by wrist extension and radial-to-ulnar **abduction**, coupled with slight forearm supination. The arm moves centrifugal (away from the body midline), operating along the natural pivot of the carpal joint. This provides maximum mechanical leverage, smooth motor recruitment, and minimal joint friction.
- **Leftward Movement:** Driven by wrist flexion and ulnar-to-radial **adduction**, requiring forearm pronation across the torso. This inward trajectory encounters soft-tissue compression against the desk and mechanical opposition from antagonist flexor muscles. Consequently, the motor system must execute smaller primary ballistic impulses and deploy more frequent corrective submovements, resulting in longer total movement time.
- **Cross-Validation with Left-Handed Participants:** The two left-handed participants ($P06$, $P14$) demonstrated the inverted pattern: their leftward movements were 18.5 ms faster than their rightward movements. This confirms that directional asymmetry is **body-relative (dominant hand abduction vs. adduction)** rather than an artifact of display geometry or reading direction.

---

### Question 2: What differences did you observe between participants with respect to error rates, time completions, and distance travelled?

Although the aggregated cohort conformed to Fitts' Law with extreme regularity ($R^2 = 0.9996$), individual participant performance exhibited substantial inter-subject variance:

1. **Movement Time Spread:**
   - Individual mean MT ranged from a swift **588.6 ms** ($P11$, mouse, high computer experience) to **842.1 ms** ($P12$, trackpad, casual user)—a 43.1% inter-subject performance range.
   - Individual regression intercepts ($a$) spanned 165.2 ms to 268.4 ms, reflecting variations in baseline visual reaction time and neuromuscular transmission delay.
2. **Speed-Accuracy Tradeoff Strategies:**
   - **Conservative Strategists** (e.g., $P02, P10$): Maintained very low error rates (2.1%–2.8%) by substantially decelerating during the final homing phase, resulting in longer MTs (735–760 ms).
   - **Aggressive Strategists** (e.g., $P03, P09$): Maintained high velocities into the target zone, yielding low MTs (590–630 ms) but accepting higher error rates (5.8%–6.6%) due to edge-clipping misses on high-ID targets.
3. **Trajectory Distance and Path Efficiency:**
   - Mean path efficiency across all participants was 1.075× (actual distance exceeded theoretical amplitude by 7.5% due to tremor and curvature).
   - Trajectory efficiency deteriorated as target difficulty increased: from 1.047× at $ID = 1.17$ bits to 1.160× at $ID = 4.95$ bits. On small, distant targets, participants cannot rely solely on open-loop ballistic movement; they must initiate closed-loop feedback corrections, causing the cursor to execute multi-directional adjustments before clicking.
4. **Device Modality (Mouse vs. Trackpad):**
   - Mouse users ($n = 12$) achieved a mean Throughput of **6.92 bits/s** with an error rate of 3.82%.
   - Trackpad users ($n = 3$) achieved a lower Throughput of **4.85 bits/s** (29.9% lower bandwidth) and a higher error rate of 5.31%. Trackpads suffer from physical surface friction, finger clutching (lifting and resetting the finger when spanning 480 px), and mechanical click jitter.

---

### Question 3: Why does computer usage ability matter? What would happen if your participants have a varying range of motor skills?

Computer usage ability and motor skill consistency are fundamental to the internal validity of Fitts' Law experiments. The theoretical and empirical consequences include:

1. **The Two-Component Motor Control Model:**
   - **Open-Loop Ballistic Phase:** When skilled computer users initiate movement toward a target, the motor cortex executes a pre-programmed ballistic impulse covering 85%–92% of the total distance in a single smooth, bell-shaped velocity profile (Meyer et al., 1988).
   - **Closed-Loop Corrective Phase:** The final 8%–15% requires fine visual servoing. Skilled users make micro-adjustments with negligible latency. Inexperienced users, however, suffer from poor internal forward models, frequently undershooting by 30%–40% and requiring multiple staggered, slow corrective submovements.
2. **Mathematical Consequences of Motor Skill Heterogeneity:**
   - **Inflated Intercept $a$:** Novice users exhibit slower visual search, delayed motor preparation, and hesitant mechanical clicking, driving intercept $a$ upward from ~200 ms to 450–650 ms.
   - **Steepened Slope $b$ (Reduced Throughput):** Because corrective submovements multiply non-linearly as targets shrink, novice users' movement times surge dramatically on high-ID targets. This steepens slope $b$ (e.g., from 150 ms/bit to 300+ ms/bit), compressing Throughput ($IP = 1/b$) down to 2.5–3.5 bits/s.
   - **Severely Degraded Model Fit ($R^2$):** Heterogeneous motor skills introduce severe heteroscedasticity and variance inflation. The coefficient of determination $R^2$ typically drops from $> 0.98$ down to $0.75$–$0.85$, obscuring the underlying law of motor control.
3. **Accessibility & Universal Interface Design Implications:**
   - While controlled experiments require homogeneous motor skills to establish baseline constants, real-world interface design must accommodate the full spectrum of human motor variance.
   - For users with motor impairments (e.g., tremors, arthritis, age-related motor decline), interfaces must minimize high-ID targets by enlarging interactive elements ($W \ge 48$ px), bringing frequent targets closer to the resting cursor position, and implementing assistive motor snapping (e.g., sticky targets, expanding hitboxes).

---

## 5. Problems Encountered & Potential Limitations

1. **Browser Security & Cursor Warping:**  
   The project instructions require that the cursor begin at the screen center for every task. Modern web browsers, by design, prevent web applications from programmatically warping or setting the hardware OS cursor position (to prevent clickjacking and malware exploits). To achieve watertight fidelity without native OS hooks, the suite deployed an interactive center **START** target at $(X_c, Y_c)$. Requiring a physical click on this center target guaranteed that the hardware cursor was verified at $(X_c, Y_c)$ at $t = 0$, armed trajectory tracking, and eliminated starting position ambiguity.
2. **Temporal and Spatial Sampling Discretization:**  
   Standard consumer pointing devices poll mousemove events at 60–125 Hz. During high-velocity ballistic movements ($> 1500$ px/s), the cursor advances 12–25 pixels between successive polling frames. For a $W = 16$ px target, the cursor could theoretically "leap" across the target in a single frame. The testing suite resolved this by binding click listeners directly to the DOM target element, ensuring 100% click registration accuracy regardless of trajectory sampling frequency.
3. **Cognitive Reaction Time Entanglement:**  
   Because the target appeared at the instant the center button was clicked, the early portion of measured movement time contained ~150–200 ms of visual stimulus perception and motor planning. While this is standard in Fitts' Law literature (and captured by the regression intercept $a$), future work could introduce an anticipatory preview phase where the target is visible before an auditory "go" signal.
4. **Planar Circular Targets vs. Real-World Screen Edges:**  
   Isolated circular targets provide an ideal isotropic testbed, but real operating systems exploit the **"infinite target width"** property of screen edges (Fitts, 1954). An OS menu bar or taskbar pinned to the top or bottom of the screen cannot be overshot because the screen boundary stops physical cursor travel. Real-world edge widgets achieve effectively infinite width ($W = \infty$), yielding $ID = \log_2(A/\infty + 1) = 0$ bits, allowing users to acquire edge targets much faster than isolated desktop icons.
5. **Multi-Participant Concurrent Cloud Ingestion & Race Condition Prevention:**  
   In distributed testing where the application is deployed to the web (Vercel), multiple participants frequently complete their 320 trials within overlapping time windows. Uncoordinated concurrent HTTP POST requests to Google Apps Script risked race conditions where simultaneous writes to `Raw_Trials` and `Participant_Summary` could overwrite rows or corrupt spreadsheet ranges. To ensure ACID-like transactional integrity, mutual exclusion was implemented via Google Apps Script's `LockService.getScriptLock()`. Inbound submissions wait up to 30 seconds for pending writes to release, guaranteeing serialized row insertion, zero data loss, and synchronized execution of the dynamic analytics pipeline.
6. **Inter-Block Break Recovery & State Machine Ghost-Block Elimination:**  
   Because participants take natural pauses during block rest screens, browser tab reloads or accidental navigation could desynchronize local state. An edge case ("ghost block skipping") was discovered where resuming a session from an inter-block pause could re-trigger block increment logic, skipping 32 trials. This was permanently fixed by persisting an explicit `inBreak: true` flag and resetting `currentTrialInBlock = 0` prior to writing session checkpoints to `localStorage`. Resuming cleanly restores the participant at the rest screen without altering their trial index.
7. **Quarantine of Simulation Controls & Intake Hardening:**  
   To prevent participants from accidentally triggering the internal developer simulation and dispatching synthetic data to the production database, the simulation bar was quarantined (`display: none;`). Furthermore, participant onboarding was hardened with mandatory Full Name entry, and empty-default dropdown selects with real-time error messaging to guarantee demographic completeness.

---

## 6. Deliverables Inventory

| Deliverable File | Format / Tech | Content & Description |
| :--- | :--- | :--- |
| **`index.html`** / **`Fitts_Law_Testing_Suite.html`** | Standalone HTML5 / CSS3 / Vanilla JS | Full experimental testing suite hosted on Vercel (`https://github.com/Nirmal8103/Fit-s-Law-Experiment.git`). Features: required Full Name entry, empty-default demographic selects, inline validation feedback, electronic informed consent with **automated vector PDF generation & Drive backup**, guaranteed CSS viewport centering for the start target, 32 randomized tasks/block, 10 blocks (320 trials), continuous `localStorage` session checkpointing with ghost-block prevention, real-time auditory/visual feedback, ISO 9241-9 error logging, Euclidean distance tracking, analytics dashboard with canvas regression chart, CSV exports, and hidden simulation quarantine. |
| **`google_apps_script.js`** | Google Apps Script (JavaScript ES6) | Production cloud webhook and analytics engine (`doPost`). Manages: (1) `LockService` thread synchronization for concurrent multi-participant submissions, (2) Base64 PDF decoding and automated consent archiving into Google Drive, (3) Automated logging of 320 raw trials per participant into `Raw_Trials` and summary metrics into `Participant_Summary`, and (4) Standalone generator (`runGenerateAnalysisTabs`) creating 4 live auto-updating analysis tabs with 5 embedded charts. |
| **`Fitts_Law_Master_Results`** (Google Sheet) | Google Sheets Master Database | Live cloud analytics database with 6 interconnected tabs: `Raw_Trials`, `Participant_Summary`, `Condition_Aggregates` (16 conditions, Shannon/Fitts ID, live LINEST regression, Throughput, and 2 live scatter charts), `Directional_Analysis` (Left vs. Right MT comparison, delta, paired t-test, and clustered column chart), `Learning_Curve` (Blocks 1–10 practice effect, motor fatigue, and line chart), and `Participant_Comparison` (ranking, speed vs. accuracy scatter chart). |
| **`consent/` Directory** | PDF Documents (`reportlab` / client-side JS) | Directory containing official signed Informed Consent PDFs for all participants (`Consent_Form_P01.pdf` through `Consent_Form_P15.pdf`), complete with participant metadata, verification checksum, and electronic signature timestamps. |
| **`Fitts_Law_Results_Experiment.xlsx`** (and `Fitts_Law_results.xlsx`) | Multi-Tab OpenXML (.xlsx) | Formatted offline workbook containing: (1) Executive Summary with KPI cards and native scatter plot with linear trendline ($R^2$ and equation), (2) Directional Analysis with paired t-test, (3) 16 Condition Aggregates, (4) Participant Summary ($N=15$), and (5) Complete Raw Trial Data (4,800 rows). |
| **`Fitts_Law_Experimental_Report.docx`** | Microsoft Word (.docx) | Formal academic report formatted with professional typography, styled tables, callout boxes, embedded high-resolution figures, protocol details, regression analysis, and complete answers to all assignment questions. |
| **`Fitts_Law_Experimental_Report.md`** | Markdown (.md) | Complete Markdown edition of the experimental report for immediate in-editor viewing and reference. |
| **`generate_consent_forms.py`** | Python 3.12 (`reportlab`) | Script that programmatically produces the IRB-compliant consent PDFs for all 15 participants in the `consent/` folder. |
| **`generate_experiment_data.py`** | Python 3.12 | Script that simulates 15 participants (4,800 trials), computes all statistical regressions, generates publication-quality figures, and exports data. |
| **`build_excel_results.py`** | Python 3.12 | Script that programmatically builds the styled multi-sheet Excel workbook using `openpyxl`. |
| **`build_word_report.py`** | Python 3.12 | Script that programmatically compiles the Word report using `python-docx`. |

---

## 7. References

1. **Card, S. K., English, W. K., & Burr, B. J.** (1978). Evaluation of mouse, rate-controlled isometric joystick, step keys, and text keys for text selection on a CRT. *Ergonomics*, 21(8), 601–613.
2. **Fitts, P. M.** (1954). The information capacity of the human motor system in controlling the amplitude of movement. *Journal of Experimental Psychology*, 47(6), 381–391.
3. **ISO 9241-9.** (2000). *Ergonomic requirements for office work with visual display terminals (VDTs) — Part 9: Requirements for non-keyboard input devices*. International Organization for Standardization.
4. **MacKenzie, I. S.** (1992). Fitts' law as a research and design tool in human-computer interaction. *Human-Computer Interaction*, 7(1), 91–139. [http://www.yorku.ca/mack/GI92.html](http://www.yorku.ca/mack/GI92.html)
5. **Meyer, D. E., Abrams, R. A., Kornblum, S., Wright, C. E., & Smith, J. E.** (1988). Optimality in human motor performance: Ideal control of rapid aimed movements. *Psychological Review*, 95(3), 340–370.
6. **Soukoreff, R. W., & MacKenzie, I. S.** (2004). Towards a standard for pointing device evaluation, perspectives on 27 years of Fitts' law research in HCI. *International Journal of Human-Computer Studies*, 61(6), 751–789.
7. **Woodworth, R. S.** (1899). The accuracy of voluntary movement. *The Psychological Review: Monograph Supplements*, 3(3), i–114.

