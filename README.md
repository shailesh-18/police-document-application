# Police Documents Generator

A comprehensive Flask web application designed for police officers to efficiently generate, manage, and download official case documents. The application supports Gujarati language throughout and provides real-time document preview functionality.

## Quick Start

1. Run `python app.py`
2. Open `http://localhost:5000`
3. Fill in case details
4. Select offence section
5. Click "Save and Proceed"
6. Preview and download documents

---

# How to Use - Police Document Generator

---

## English

### Step-by-Step Guide

**1. Start on Home Page**
   - You will see a form with multiple fields for case details
   - You can fill all fields or only a few fields as you go

**2. Fill in the Information**
   - Enter accused name, age, father's name
   - Enter permanent address (select district, taluka, address)
   - Enter current address (select district, taluka, address)
   - Enter mobile numbers and other personal details
   - Enter offence details, FIR number, dates, times as needed
   - All fields are optional - you don't need to fill everything

**3. Select Offence Section**
   - On the right sidebar, select the offence section (281, 302, 379, or Other)
   - This determines which documents will be generated

**4. Save and Proceed**
   - Click the "Save and Proceed" button on the right sidebar
   - If any important fields are missing, the system will show which fields need to be filled
   - Fill those missing fields and try again
   - If you have filled enough information, you will be taken to the Documents page

**5. On Documents Page - View Previews**
   - You can see all available documents for your selected offence section
   - Click the "Preview" button on any document to see how it will look with your data
   - View the merged preview panel on the right side showing all documents combined

**6. Download Options**
   - **Download Single Document**: Click download button on individual document to get DOCX file
   - **Download All as ZIP**: Download all documents in a compressed folder
   - **Download Merged Document**: Download single DOCX with all documents merged together

**7. Edit Your Information**
   - Click the "Back to Edit" button to return to the form
   - Make changes to your information
   - Select the same offence section again
   - Click "Save and Proceed" to go back to Documents page with updated information

**8. Repeat as Needed**
   - You can go back and forth between the form and documents page as many times as needed
   - Your form data is automatically saved in your browser
   - The browser will remember your entries even if you close and reopen the page

---

## ગુજરાતી

### ચરણ-દર-ચરણ માર્ગદર્શન

**1. હોમ પેજ પર શરૂ કરો**
   - તમને કેસ વિગતો માટે બહુ બધા ક્ષેત્રો સાથે એક ફોર્મ દેખાશે
   - તમે બધા ક્ષેત્રો ભરી શકો અથવા જતાં જતાં માત્ર કેટલાક ક્ષેત્રો ભરી શકો

**2. માહિતી ભરો**
   - આરોપી નું નામ, ઉંમર, પિતાનું નામ દાખલ કરો
   - કાયમી સરનામું દાખલ કરો (જિલ્લો, તાલુકો, સરનામું પસંદ કરો)
   - વર્તમાન સરનામું દાખલ કરો (જિલ્લો, તાલુકો, સરનામું પસંદ કરો)
   - મોબાઇલ નંબર અને અન્ય વ્યક્તિગત વિગતો દાખલ કરો
   - અપરાધની વિગતો, FIR નંબર, તારીખો, સમય જરૂર મુજબ દાખલ કરો
   - બધા ક્ષેત્રો વૈકલ્પિક છે - તમે બધું ભરવું જરૂરી નથી

**3. અપરાધ વિભાગ પસંદ કરો**
   - જમણા સાઇડબારમાં, અપરાધ વિભાગ પસંદ કરો (281, 302, 379, અથવા અન્ય)
   - આ તમે કઈ દસ્તાવેજો જનરેટ કરશો તે નક્કી કરે છે

**4. સંરક્ષણ અને આગળ વધો**
   - જમણા સાઇડબારમાં "સંરક્ષણ અને આગળ વધો" બટન પર ક્લિક કરો
   - જો કોઈ મહત્વપૂર્ણ ક્ષેત્રો ખૂટી રહ્યાં હોય, તો સિસ્ટમ બતાવશે કે કયા ક્ષેત્રો ભરવું પડશે
   - તે ખૂટી રહેલા ક્ષેત્રો ભરો અને ફરી પ્રયાસ કરો
   - જો તમે પર્યાપ્ત માહિતી ભર્યું હોય, તો તમે દસ્તાવેજો પેજ પર જશો

**5. દસ્તાવેજો પેજ પર - પૂર્વાવલોકન જુઓ**
   - તમે તમારા પસંદ કરેલ અપરાધ વિભાગ માટે તમામ ઉપલબ્ધ દસ્તાવેજો જોઈ શકો છો
   - કોઈ પણ દસ્તાવેજ પર "પૂર્વાવલોકન" બટન પર ક્લિક કરો તેને તમારી માહિતી સાથે કેવો દેખાશે તે જોવા માટે
   - જમણા બાજુના પૂર્વાવલોકન પેનલમાં બધી દસ્તાવેજો સાથે જોઈ શકો છો

**6. ડાઉનલોડ વિકલ્પો**
   - **એક દસ્તાવેજ ડાઉનલોડ કરો**: વ્યક્તિગત દસ્તાવેજ પર ડાઉનલોડ બટન પર ક્લિક કરો DOCX ફાઇલ મેળવવા માટે
   - **બધું ZIP તરીકે ડાઉનલોડ કરો**: બધી દસ્તાવેજો સંકુચિત ફોલ્ડરમાં ડાઉનલોડ કરો
   - **મર્જ કરેલ દસ્તાવેજ ડાઉનલોડ કરો**: બધી દસ્તાવેજો સાથે એક DOCX ફાઇલ ડાઉનલોડ કરો

**7. તમારી માહિતી સંપાદિત કરો**
   - ફોર્મ પર પાછા ફરવા માટે "સંપાદન માટે પાછા ફરો" બટન પર ક્લિક કરો
   - તમારી માહિતીમાં પરિવર્તન કરો
   - આવી જ અપરાધ વિભાગ ફરી પસંદ કરો
   - અપડેટ કરેલ માહિતી સાથે દસ્તાવેજો પેજ પર જવા માટે "સંરક્ષણ અને આગળ વધો" પર ક્લિક કરો

**8. જરૂર મુજબ પુનરાવર્તન કરો**
   - તમે ફોર્મ અને દસ્તાવેજો પેજ વચ્ચે જાણે તેટલી વાર આગળ-પાછળ જઈ શકો છો
   - તમારી ફોર્મ માહિતી આપોઆપ તમારા બ્રાઉઝરમાં સંરક્ષિત થાય છે
   - બ્રાઉઝર પૃષ્ઠ બંધ કર્યા પછી પણ તમારી પ્રવિષ્ટિઓ યાદ રાખશે અને ફરી ખોલવામાં આવશે

---


✓ Fill all or only some fields / બધા અથવા માત્ર કેટલાક ક્ષેત્રો ભરો
✓ Auto-save to browser / બ્રાઉઝરમાં આપોઆપ સંરક્ષણ
✓ Missing field validation / ખૂટી રહેલા ક્ષેત્રો ચકાસણી
✓ Multiple offence sections / બહુવિધ અપરાધ વિભાગો
✓ Live previews / લાઇવ પૂર્વાવલોકન
✓ Multiple download formats / બહુવિધ ડાઉનલોડ ફોર્મેટ
✓ Edit and return / સંપાદિત કરો અને પાછા ફરો

---

##  Technology Stack

| Component | Technology |
|-----------|-----------|
| Backend Framework | Flask (Python) |
| Frontend | HTML5, Bootstrap 5, JavaScript |
| Document Processing | python-docx, docxcompose |
| Language | Python 3.12+ |
| Deployment | Vercel |
| Data Storage | Browser localStorage + Session |

---

##  Project Structure

```
Police Document generators/
├── app.py                      # Main Flask application
├── requirements.txt            # Python dependencies
├── README.md                   # Documentation
├── vercel.json                 # Vercel deployment config
├── templates/
│   ├── index.html              # Home page (data entry form)
│   └── documents.html          # Documents page (preview & download)
├── word_templates/             # Word document templates
│   ├── 281/
│   ├── 302/
│   ├── 379/
│   └── GENERAL/
└── static/                     # Static files (JS, fonts)
```

---

##  Installation & Setup

### Local Development

1. **Clone/Download Project**
   ```bash
   cd "g:\Police Document generators"
   ```

2. **Create Virtual Environment**
   ```bash
   python -m venv py_env
   py_env\Scripts\activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run Application**
   ```bash
   python app.py
   ```
   
   App available at: `http://localhost:5000`

### Vercel Deployment

```bash
npm install -g vercel
vercel login
vercel
```

---

## Required Fields for Each Offence Section

### Section 281
- Accused name, age, address
- Offence date and time
- Arrest date and time

### Section 302
- All accused details
- Complete case information
- Offence and arrest details
- FIR number

### Section 379
- Accused information
- Property details
- Crime location and date
- Recovery information

### Section Other (GENERAL)
- Basic accused details
- Case reference number
- Offence description

---

## ➕ Adding New Templates

### To Add Template to Existing Section

1. Create Word document with placeholders in format: `[field_name]`
2. Save file in correct folder: `word_templates/{section}/`
3. Update `OFFENCE_MAPPING` in `app.py`
4. Restart application

**Example:**
```python
OFFENCE_MAPPING = {
    "281": ["Arrest_memo_281.docx", "Bail_Bond_281.docx", "Notice_281.docx"],
    "302": ["Arrest_Memo_Major.docx", ...],
    ...
}
```

### To Add New Offence Section

1. Create folder: `word_templates/{section_number}/`
2. Add template files to folder
3. Update `OFFENCE_MAPPING` in `app.py`
4. Update offence_section dropdown in `index.html`
5. Restart application

---

## 📋 DETAILED GUIDE: Adding New Offence Sections

This guide provides step-by-step instructions for adding a new offence section to the Police Document Generator system.

### Overview of System Architecture

The system uses a **4-layer architecture** for offence sections:

```
User Selection (index.html dropdown)
         ↓
Backend Mapping (app.py OFFENCE_MAPPING)
         ↓
Template Loading (word_templates/{section}/)
         ↓
Preview Generation (word_previews/{preview_file}.html)
```

---

### Step 1: Create Template Documents

**What:** Create Word documents with placeholder fields that will be replaced with user data

**Where:** `word_templates/{section_number}/` folder

**Files to Create:** 
- One or more `.docx` files named descriptively
- Example: `Arrest_Memo_370.docx`, `FIR_Notice_370.docx`, etc.

**Format:**
1. Open Microsoft Word or any compatible editor
2. Create your document with the official format/template
3. Where dynamic data should appear, insert placeholders in format: `[field_name]`
4. Save as `.docx` format (not `.doc`)

**Available Field Names:**
Use any of these field names as placeholders. They will be replaced with actual user data:

#### Accused Information
```
[acc_name]              - Accused name
[acc_father]            - Father's name
[acc_surname]           - Surname/Last name
[acc_alias]             - Alias/उर्फे
[acc_gender]            - Gender (Male/Female)
[acc_age]               - Age
[acc_dob]               - Date of birth
[acc_religion]          - Religion
[acc_caste]             - Caste
[acc_subcaste]          - Subcaste
[acc_nationality]       - Nationality
[acc_marital]           - Marital status
```

#### Addresses
```
[perm_house]            - House number (Permanent)
[perm_area]             - Area (Permanent)
[perm_village]          - Village (Permanent)
[perm_district]         - District (Permanent)
[perm_taluka]           - Taluka (Permanent)
[perm_state]            - State (Permanent)
[perm_pin]              - Pincode (Permanent)

[curr_address]          - Full current address
[curr_city]             - City (Current)
[curr_district]         - District (Current)
[curr_taluka]           - Taluka (Current)
[curr_state]            - State (Current)
[curr_pin]              - Pincode (Current)
```

#### Occupation
```
[occ_type]              - Type of occupation
[occ_place]             - Place of work
[occ_income]            - Income
```

#### Contact & ID
```
[mobile_1]              - Primary mobile number
[mobile_2]              - Alternate mobile number
[id_type]               - ID type (Aadhar, DL, etc.)
[id_number]             - ID number
```

#### Physical Description
```
[phy_height]            - Height
[phy_build]             - Build/Body type
[phy_complexion]        - Complexion
[phy_eyes]              - Eye color
[phy_hair]              - Hair color
[phy_facial_hair]       - Beard/Mustache
```

#### Marks & Identification
```
[mark_1]                - Identification mark 1
[mark_2]                - Identification mark 2
[old_wounds]            - Old wounds
[other_id_marks]        - Other marks/scars
```

#### Relative Information
```
[rel_name]              - Relative's name
[rel_relation]          - Relation to accused
[rel_mobile]            - Relative's mobile
[rel_address]           - Relative's address
```

#### Case Details
```
[case_ps]               - Police station
[case_district]         - District
[case_taluka]           - Taluka
[crime_no]              - Crime/FIR number
[crime_type]            - Type of crime
[crime_year]            - Year of crime
[offence_desc]          - Offence description
[offence_section]       - Applicable section/law
[offence_place]         - Place of offence
[offence_date]          - Date of offence
[offence_time]          - Time of offence
```

#### Arrest Information
```
[is_arrested]           - Whether arrested (Yes/No)
[arrest_date]           - Arrest date
[arrest_time]           - Arrest time
[arrest_place]          - Place of arrest
[arrest_entry_no]       - Station diary entry number
```

#### Bail Information
```
[is_bailed]             - Whether on bail
[bail_authority]        - Authority granting bail
[bail_date]             - Bail date
[bail_time]             - Bail time
[bail_conditions]       - Bail conditions
[bail_person_name]      - Bail person's name
```

#### Status & Intimation
```
[status_proven]         - Proven status
[status_chargesheet]    - Chargesheet status
[status_case]           - Case status
[status_release_date]   - Release date
[intimation_method]     - Method of intimation
[intimation_date]       - Intimation date
[intimation_time]       - Intimation time
[intimation_entry_no]   - Intimation entry number
[intimation_entry_time] - Intimation time entry
```

#### Officer Information
```
[io_name]               - Investigating officer name
[io_designation]        - IO's designation
[io_buckle]             - IO's buckle/badge number
[io_police_station]     - IO's police station
[remarks]               - Remarks/Notes
```

#### Complainant Information
```
[complain_person]       - Complainant's name
[complain_person_work]  - Complainant's work/PS
[comp_per_mobile]       - Complainant's mobile
[comp_per_address]      - Complainant's address
```

#### Panchnama (Witness Documentation)
```
[panchnumu_date]        - Panchnama date
[panchnamu_time]        - Panchnama time
[panch_clerk]           - Panchnama clerk name
[panch1_name]           - First panch's name
[panch1_surname]        - First panch's surname
[panch1_age]            - First panch's age
[panch1_work]           - First panch's occupation
[panch1_address]        - First panch's address
[panch2_name]           - Second panch's name
[panch2_surname]        - Second panch's surname
[panch2_age]            - Second panch's age
[panch2_work]           - Second panch's occupation
[panch2_address]        - Second panch's address
```

#### Administrative
```
[pso_name]              - PSO name
[court_place]           - Court location
[ro_name]               - Review officer's name
[ro_designation]        - RO's designation
[ro_buckle]             - RO's buckle number
[ro_police_station]     - RO's police station
[auth_date]             - Authentication date
[auth_place]            - Authentication place
[auth_print_name]       - Name for signature
```

**Example Word Document:**
```
ARREST MEMORANDUM - SECTION [offence_section]

Date: [offence_date]                    Time: [offence_time]
Location: [offence_place]

ACCUSED DETAILS:
Name: [acc_name]
Father's Name: [acc_father]
Surname: [acc_surname]
Age: [acc_age] years
Gender: [acc_gender]

PERMANENT ADDRESS:
[perm_house], [perm_area]
[perm_village], [perm_district]
[perm_taluka], [perm_state] - [perm_pin]

CURRENT ADDRESS:
[curr_address]

CONTACT:
Mobile: [mobile_1]
Alternate: [mobile_2]

OFFENSE DETAILS:
Description: [offence_desc]
Crime Number: [crime_no]
FIR Year: [crime_year]

ARREST DETAILS:
Arrest Date: [arrest_date]
Arrest Time: [arrest_time]
Place: [arrest_place]

INVESTIGATING OFFICER:
Name: [io_name]
Designation: [io_designation]
Police Station: [io_police_station]
Buckle No: [io_buckle]

Signature: ________________     Date: [auth_date]
           [auth_print_name]
```

---

### Step 2: Create Directory Structure

**What:** Create the folder for your new section in templates directory

**Where:** In the project root directory

**Commands:**
```bash
# Navigate to project root
cd "g:\Police Document generators"

# Create new section folder (replace 370 with your section number)
mkdir word_templates\370

# Your files should be placed here:
# word_templates\370\Arrest_Memo_370.docx
# word_templates\370\FIR_Notice_370.docx
# etc.
```

**Result Structure:**
```
word_templates/
├── 281/
│   ├── Arrest_memo_281.docx
│   ├── Bail_Bond_281.docx
│   └── Notice_281.docx
├── 302/
│   ├── Arrest_Memo_Major.docx
│   ├── Remand_Application.docx
│   └── Panchnama_Scene.docx
├── 66_1B/
│   └── kalam66_b1.docx
├── 370/              ← NEW SECTION
│   ├── Arrest_Memo_370.docx
│   ├── FIR_Notice_370.docx
│   └── Bail_Bond_370.docx
└── GENERAL/
    └── Standard_Intimation.docx
```

---

### Step 3: Update Backend Mapping (app.py)

**What:** Register your new section in the backend so the system knows which templates belong to it

**Where:** `app.py` file, around **line 40-60**

**Current Code:**
```python
OFFENCE_MAPPING = {
    "281": [
        "Arrest_memo_281.docx",
        "Bail_Bond_281.docx",
        "Notice_281.docx"
    ],
    "302": [
        "Arrest_Memo_Major.docx",
        "Remand_Application.docx",
        "Panchnama_Scene.docx"
    ],
    "66_1B": [
        "kalam66_b1.docx"
    ],
    "GENERAL": [
        "Standard_Intimation.docx"
    ]
}
```

**What to Add:**
Add your new section with all its template files:

```python
OFFENCE_MAPPING = {
    "281": [
        "Arrest_memo_281.docx",
        "Bail_Bond_281.docx",
        "Notice_281.docx"
    ],
    "302": [
        "Arrest_Memo_Major.docx",
        "Remand_Application.docx",
        "Panchnama_Scene.docx"
    ],
    "66_1B": [
        "kalam66_b1.docx"
    ],
    "370": [                               ← ADD THIS NEW SECTION
        "Arrest_Memo_370.docx",
        "FIR_Notice_370.docx",
        "Bail_Bond_370.docx"
    ],
    "GENERAL": [
        "Standard_Intimation.docx"
    ]
}
```

**Rules:**
- Section name (e.g., "370") must match your folder name in `word_templates`
- List all `.docx` files in that section's folder
- Use exact filenames as they appear in folder
- File names are case-sensitive

---

### Step 4: Create HTML Preview File

**What:** Create an HTML preview file that shows how the document looks in real-time as user edits form

**Where:** `word_previews/` folder (if creating a preview for the main form document)

**Recommendation:** Create preview HTML for the main/primary template of your section

**File Naming:** Match the template name but with `.html` extension
- Example: `Arrest_Memo_370.docx` → `Arrest_Memo_370.html`

**For Section 370 - Example Preview HTML:**

Create file: `word_previews/Arrest_Memo_370.html`

```html
<!DOCTYPE html>
<html lang="gu">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Arrest Memo - Section 370</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Arial', sans-serif;
            background-color: #f0f0f0;
            padding: 20px;
        }

        .document {
            width: 794px;
            height: auto;
            margin: 0 auto;
            background: white;
            padding: 40px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            line-height: 1.6;
        }

        .variable {
            border-bottom: 1px dotted #999;
            min-width: 100px;
            display: inline-block;
            cursor: text;
            padding: 2px 4px;
            border-radius: 2px;
        }

        .variable:hover {
            background-color: #fffacd;
        }

        .variable:focus {
            background-color: #ffff99;
            outline: none;
        }

        h2 {
            text-align: center;
            margin-bottom: 30px;
            font-size: 18px;
            font-weight: bold;
        }

        .section {
            margin-bottom: 20px;
        }

        .section-title {
            font-weight: bold;
            background-color: #e0e0e0;
            padding: 5px;
            margin-bottom: 10px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 15px;
        }

        td {
            padding: 5px;
            border: 1px solid #ccc;
        }

        .label {
            font-weight: bold;
            width: 30%;
            background-color: #f5f5f5;
        }

        .value {
            width: 70%;
        }
    </style>
</head>
<body>

<div class="document">
    <h2>ARREST MEMORANDUM - SECTION 370</h2>
    
    <div class="section">
        <div class="section-title">DATE & TIME OF OFFENSE</div>
        <table>
            <tr>
                <td class="label">Date:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('offence_date', this.textContent)">{{ getVariableValue('offence_date') }}</span></td>
                <td class="label">Time:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('offence_time', this.textContent)">{{ getVariableValue('offence_time') }}</span></td>
            </tr>
            <tr>
                <td class="label">Location:</td>
                <td colspan="3" class="value"><span class="variable" contenteditable onblur="updateVariable('offence_place', this.textContent)">{{ getVariableValue('offence_place') }}</span></td>
            </tr>
        </table>
    </div>

    <div class="section">
        <div class="section-title">ACCUSED DETAILS</div>
        <table>
            <tr>
                <td class="label">Name:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('acc_name', this.textContent)">{{ getVariableValue('acc_name') }}</span></td>
                <td class="label">Father's Name:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('acc_father', this.textContent)">{{ getVariableValue('acc_father') }}</span></td>
            </tr>
            <tr>
                <td class="label">Surname:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('acc_surname', this.textContent)">{{ getVariableValue('acc_surname') }}</span></td>
                <td class="label">Age:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('acc_age', this.textContent)">{{ getVariableValue('acc_age') }}</span></td>
            </tr>
            <tr>
                <td class="label">Gender:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('acc_gender', this.textContent)">{{ getVariableValue('acc_gender') }}</span></td>
                <td class="label">Mobile:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('mobile_1', this.textContent)">{{ getVariableValue('mobile_1') }}</span></td>
            </tr>
        </table>
    </div>

    <div class="section">
        <div class="section-title">PERMANENT ADDRESS</div>
        <table>
            <tr>
                <td class="label">Address:</td>
                <td colspan="3" class="value">
                    <span class="variable" contenteditable onblur="updateVariable('perm_house', this.textContent)">{{ getVariableValue('perm_house') }}</span>,
                    <span class="variable" contenteditable onblur="updateVariable('perm_area', this.textContent)">{{ getVariableValue('perm_area') }}</span>,
                    <span class="variable" contenteditable onblur="updateVariable('perm_village', this.textContent)">{{ getVariableValue('perm_village') }}</span>
                </td>
            </tr>
            <tr>
                <td class="label">District:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('perm_district', this.textContent)">{{ getVariableValue('perm_district') }}</span></td>
                <td class="label">Taluka:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('perm_taluka', this.textContent)">{{ getVariableValue('perm_taluka') }}</span></td>
            </tr>
            <tr>
                <td class="label">Pincode:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('perm_pin', this.textContent)">{{ getVariableValue('perm_pin') }}</span></td>
                <td class="label">State:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('perm_state', this.textContent)">{{ getVariableValue('perm_state') }}</span></td>
            </tr>
        </table>
    </div>

    <div class="section">
        <div class="section-title">OFFENSE DETAILS</div>
        <table>
            <tr>
                <td class="label">Section:</td>
                <td colspan="3" class="value"><span class="variable" contenteditable onblur="updateVariable('offence_section', this.textContent)">{{ getVariableValue('offence_section') }}</span></td>
            </tr>
            <tr>
                <td class="label">Description:</td>
                <td colspan="3" class="value"><span class="variable" contenteditable onblur="updateVariable('offence_desc', this.textContent)">{{ getVariableValue('offence_desc') }}</span></td>
            </tr>
            <tr>
                <td class="label">FIR Number:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('crime_no', this.textContent)">{{ getVariableValue('crime_no') }}</span></td>
                <td class="label">Police Station:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('case_ps', this.textContent)">{{ getVariableValue('case_ps') }}</span></td>
            </tr>
        </table>
    </div>

    <div class="section">
        <div class="section-title">ARREST INFORMATION</div>
        <table>
            <tr>
                <td class="label">Arrest Date:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('arrest_date', this.textContent)">{{ getVariableValue('arrest_date') }}</span></td>
                <td class="label">Arrest Time:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('arrest_time', this.textContent)">{{ getVariableValue('arrest_time') }}</span></td>
            </tr>
            <tr>
                <td class="label">Place of Arrest:</td>
                <td colspan="3" class="value"><span class="variable" contenteditable onblur="updateVariable('arrest_place', this.textContent)">{{ getVariableValue('arrest_place') }}</span></td>
            </tr>
        </table>
    </div>

    <div class="section">
        <div class="section-title">INVESTIGATING OFFICER</div>
        <table>
            <tr>
                <td class="label">IO Name:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('io_name', this.textContent)">{{ getVariableValue('io_name') }}</span></td>
                <td class="label">Designation:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('io_designation', this.textContent)">{{ getVariableValue('io_designation') }}</span></td>
            </tr>
            <tr>
                <td class="label">Police Station:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('io_police_station', this.textContent)">{{ getVariableValue('io_police_station') }}</span></td>
                <td class="label">Buckle No:</td>
                <td class="value"><span class="variable" contenteditable onblur="updateVariable('io_buckle', this.textContent)">{{ getVariableValue('io_buckle') }}</span></td>
            </tr>
        </table>
    </div>

    <div class="section" style="margin-top: 40px; text-align: center;">
        <p style="margin-bottom: 30px;">Authorized Officer Signature</p>
        <p>________________</p>
        <p><span class="variable" contenteditable onblur="updateVariable('auth_print_name', this.textContent)">{{ getVariableValue('auth_print_name') }}</span></p>
        <p><span class="variable" contenteditable onblur="updateVariable('auth_date', this.textContent)">{{ getVariableValue('auth_date') }}</span></p>
    </div>
</div>

<script>
    // Get variable value from localStorage
    function getVariableValue(key) {
        try {
            const caseData = JSON.parse(localStorage.getItem('case_data') || '{}');
            return caseData[key] || '';
        } catch (e) {
            return '';
        }
    }

    // Update variable in localStorage
    function updateVariable(key, newValue) {
        try {
            const caseData = JSON.parse(localStorage.getItem('case_data') || '{}');
            caseData[key] = newValue;
            localStorage.setItem('case_data', JSON.stringify(caseData));
            
            // Notify parent window of update
            if (window.parent !== window) {
                window.parent.postMessage({
                    type: 'updateField',
                    key: key,
                    value: newValue
                }, '*');
            }
        } catch (e) {
            console.error('Error updating variable:', e);
        }
    }

    // Listen for external changes to localStorage
    window.addEventListener('storage', (e) => {
        if (e.key === 'case_data' && e.newValue) {
            location.reload();
        }
    });
</script>

</body>
</html>
```

**How to Create Your Own Preview:**
1. Use the above as template
2. Replace section name "370" with your section
3. Add/remove rows based on fields you want to show
4. Keep the `contenteditable` spans for real-time editing
5. Ensure `getVariableValue()` and `updateVariable()` calls use correct field names

**Alternative (Simpler):** You can also just duplicate and modify an existing preview file (e.g., `Arrest_memo_281.html`)

---

### Step 5: Update Frontend Dropdown (index.html)

**What:** Add your new section to the offence section dropdown so users can select it

**Where:** `templates/index.html` file, around **line 440-445**

**Current Code:**
```html
<select name="offence_section" class="form-control" style="border-color: #ffc107;" onchange="updatePreview('offence_section', this.value)">
    <!-- <option value="281" {% if data.offence_section == '281' %}selected{% endif %}>281 (BNS)</option>
    <option value="302" {% if data.offence_section == '302' %}selected{% endif %}>302 (Murder)</option> -->
    <option value="66_1B" {% if data.offence_section == '66_1B' %}selected{% endif %}>66(1)B</option>
    <!-- <option value="Other" {% if data.offence_section == 'Other' %}selected{% endif %}>Other</option> -->
</select>
```

**What to Add:**
Uncomment existing or add your new section:

```html
<select name="offence_section" class="form-control" style="border-color: #ffc107;" onchange="updatePreview('offence_section', this.value)">
    <option value="281" {% if data.offence_section == '281' %}selected{% endif %}>281 (BNS)</option>
    <option value="302" {% if data.offence_section == '302' %}selected{% endif %}>302 (Murder)</option>
    <option value="370" {% if data.offence_section == '370' %}selected{% endif %}>370 (Human Trafficking)</option>  ← ADD THIS
    <option value="66_1B" {% if data.offence_section == '66_1B' %}selected{% endif %}>66(1)B</option>
    <option value="Other" {% if data.offence_section == 'Other' %}selected{% endif %}>Other</option>
</select>
```

**Rules:**
- `value` must match the key in `OFFENCE_MAPPING` (e.g., "370")
- Display text can be anything descriptive (e.g., "370 (Human Trafficking)")
- Maintain the Jinja2 syntax: `{% if data.offence_section == '370' %}selected{% endif %}`

---

### Step 6: Update documents.html Preview Mapping

**What:** Map the section to its preview file so the correct preview loads when selected

**Where:** `templates/documents.html` file, around **line 540-555**

**Current Code:**
```javascript
function loadMergedPreview() {
    const container = document.getElementById('mergedPreview');
    const caseData = JSON.parse(localStorage.getItem('case_data') || '{}');
    const offenceSection = caseData.offence_section || '281';
    
    const previewMap = {
        '281': 'Arrest_memo_281.html',
        '66_1B': 'kalam66_b1.html'
    };
    
    const previewFile = previewMap[offenceSection] || 'kalam66_b1.html';
    container.innerHTML = '<iframe src="/word_previews/' + previewFile + '" style="width: 100%; height: 100%; border: none; border-radius: 6px;"></iframe>';
}
```

**What to Add:**
Add your section to the `previewMap`:

```javascript
function loadMergedPreview() {
    const container = document.getElementById('mergedPreview');
    const caseData = JSON.parse(localStorage.getItem('case_data') || '{}');
    const offenceSection = caseData.offence_section || '281';
    
    const previewMap = {
        '281': 'Arrest_memo_281.html',
        '66_1B': 'kalam66_b1.html',
        '370': 'Arrest_Memo_370.html'    ← ADD THIS LINE
    };
    
    const previewFile = previewMap[offenceSection] || 'kalam66_b1.html';
    container.innerHTML = '<iframe src="/word_previews/' + previewFile + '" style="width: 100%; height: 100%; border: none; border-radius: 6px;"></iframe>';
}
```

**Rules:**
- Key must match section name (e.g., '370')
- Value must match HTML preview filename in `word_previews/` folder
- Update will automatically reload correct preview when user changes section

---

### Step 7: Restart Application

**What:** Restart Flask server for changes to take effect

**Commands:**
```bash
# Stop current server (Ctrl + C in terminal)
Ctrl + C

# Restart server
python app.py

# Check output:
# * Running on http://127.0.0.1:5000
# * WARNING: This is a development server...
```

**Verification:**
1. Open `http://localhost:5000`
2. Check offence section dropdown - should show new section
3. Select new section and click "Save and Proceed"
4. Should see templates for new section in documents page
5. Click preview - should show your HTML preview

---

### Step 8: Test the New Section

**Testing Checklist:**

```
✓ Form Page (index.html):
  □ New section appears in offence dropdown
  □ Can select new section
  □ Form fields are visible and editable
  □ Data saves when clicking inputs
  
✓ Documents Page (documents.html):
  □ New section's templates appear in list
  □ Can preview each template
  □ Real-time preview updates as form changes
  □ Can download individual documents
  □ Can download merged document
  
✓ Data Persistence:
  □ Reload page - data still there
  □ Go to documents page and back to form - data preserved
  □ Edit in preview - updates appear in form
  □ Download captures latest data
  
✓ Multiple Documents:
  □ Each template in section downloads independently
  □ ZIP download includes all templates
  □ Merged download combines all templates
```

---

### Complete Example: Adding Section 370 (Human Trafficking)

**Summary of all files to create/modify:**

1. **Create Folder:** `word_templates/370/`

2. **Create Template:** `word_templates/370/Arrest_Memo_370.docx`
   - Format official arrest memo for Section 370
   - Use placeholders: [acc_name], [offence_date], etc.

3. **Create Preview:** `word_previews/Arrest_Memo_370.html`
   - Copy template structure from existing preview
   - Add contenteditable fields

4. **Update:** `app.py` line ~40
   ```python
   "370": ["Arrest_Memo_370.docx"]
   ```

5. **Update:** `templates/index.html` line ~440
   ```html
   <option value="370" ...>370 (Human Trafficking)</option>
   ```

6. **Update:** `templates/documents.html` line ~545
   ```javascript
   '370': 'Arrest_Memo_370.html',
   ```

7. **Restart:** `python app.py`

8. **Test:** Form → Save → Documents → Preview/Download

---

### Troubleshooting New Sections

| Problem | Solution |
|---------|----------|
| New section doesn't appear in dropdown | Restart Flask server, clear browser cache |
| Templates not found | Check folder path matches section name exactly (case-sensitive) |
| Documents won't load | Verify `.docx` filenames match `OFFENCE_MAPPING` exactly |
| Preview is blank | Check HTML preview file syntax, verify localStorage has data |
| Data not syncing to preview | Ensure preview HTML has `getVariableValue()` and `updateVariable()` functions |
| Download creates empty document | Check placeholders in Word document use exact format: `[field_name]` |
| Field data not replacing in document | Verify field name in placeholder matches `REQUIRED_FIELDS` list in app.py |

---

### Adding Fields to New Section

If you need to add NEW FIELDS beyond the predefined list:

1. **Add to `app.py` REQUIRED_FIELDS list** (~line 120)
   ```python
   REQUIRED_FIELDS = [
       # ... existing fields ...
       "new_field_name"  ← ADD HERE
   ]
   ```

2. **Add to `app.py` FIELD_LABELS dictionary** (~line 230)
   ```python
   FIELD_LABELS = {
       # ... existing labels ...
       "new_field_name": "ફીલ્ડ લેબલ (Field Label)"  ← ADD HERE
   }
   ```

3. **Add to `templates/index.html` form** (~line 400+)
   ```html
   <label>Field Label</label>
   <input type="text" name="new_field_name" class="form-control" value="{{ data.new_field_name }}">
   ```

4. **Use in Word template:**
   ```
   [new_field_name]
   ```

5. **Use in HTML preview:**
   ```html
   <span class="variable" contenteditable onblur="updateVariable('new_field_name', this.textContent)">{{ getVariableValue('new_field_name') }}</span>
   ```

---

### Best Practices

✅ **DO:**
- Use descriptive section names (section number recommended)
- Create preview HTML for better user experience
- Test all templates before deployment
- Document any special field requirements
- Use consistent placeholder formatting: `[field_name]`
- Keep template filenames descriptive

❌ **DON'T:**
- Use spaces in section names (breaks mapping)
- Use special characters in filenames
- Create circular template dependencies
- Skip updating documentation
- Mix old and new placeholder formats
- Leave placeholder placeholders in documents

---

**Last Updated:** February 7, 2026

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Placeholders not replacing | Check format is `[field_name]`, verify field names match |
| Template not found | Verify file path and filename (case-sensitive) |
| Data not saving | Check browser storage is enabled, clear cache |
| Documents not loading | Restart Flask app, verify template files exist |

---

## Notes

- All form fields are optional
- Data auto-saves to browser localStorage
- Browser remembers entries when reopened
- Templates use `[field_name]` placeholder format
- Offence section dropdown controls available templates

---

**Last Updated:** January 31, 2026
