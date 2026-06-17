# Local Storage & Field Management Analysis

## Overview
Your project uses a **dual-layer data storage system** combining **browser localStorage** (frontend) with **Flask Session** (backend).

---

## 1. HOW FIELDS ARE STORED

### **Frontend Storage (Browser LocalStorage)**
**File:** `templates/index.html` (Lines 273-295)

```javascript
// Save all form fields to browser localStorage
function saveToLocalStorage() {
    let data = {};
    document.querySelectorAll('input, select, textarea').forEach(el => {
        if (el.name) data[el.name] = el.value;
    });
    localStorage.setItem('case_data', JSON.stringify(data));
    
    // Also sync to backend
    fetch('/api/save_data', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(data)
    });
}
```

**Storage Details:**
- **Key:** `'case_data'`
- **Format:** JSON string containing all form fields
- **Trigger:** Fires on any `input` or `change` event
- **Location:** Browser's localStorage (persists until cleared)

### **Backend Storage (Flask Session)**
**File:** `app.py` (Lines 248-260)

```python
def load_data():
    if 'case_data' in session:
        return session['case_data']
    return {}

def save_data(data):
    session['case_data'] = data
    session.modified = True
```

**Storage Details:**
- **Location:** Server-side Flask session
- **Persistence:** In-memory (or configured session backend)
- **Key:** `'case_data'`
- **Automatic Sync:** Triggered by API endpoint `/api/save_data`

---

## 2. HOW FIELDS ARE ACCESSED ON OTHER FILES

### **Method 1: Direct localStorage Access (Client-Side)**

#### On `index.html` (Main Form Page):
```javascript
// Load saved data when page loads
function loadFromLocalStorage() {
    let saved = localStorage.getItem('case_data');
    if (saved) {
        let data = JSON.parse(saved);
        for (let key in data) {
            let input = document.querySelector(`[name="${key}"]`);
            if (input) {
                input.value = data[key];
                updatePreview(key, data[key]);
            }
        }
    }
}

// Initialize on page load
window.addEventListener('load', () => {
    loadFromLocalStorage();
});
```

#### On `documents.html` (Preview/Generation Page):
The preview pages access localStorage like this:
```javascript
// Get variable value from localStorage
const value = localStorage.getItem(key);
```

**File:** `word_previews/kalam66_b1.html` (Lines 323-325)

---

### **Method 2: Server-Side Access (Backend)**

#### API Endpoints for Data Retrieval:

**1. GET endpoint (Retrieve data from session):**
```python
@app.route('/api/load_data', methods=['GET'])
def api_load_data():
    data = load_data()  # Gets from session['case_data']
    return data, 200
```

**2. POST endpoint (Store data to session):**
```python
@app.route('/api/save_data', methods=['POST'])
def api_save_data():
    data = request.get_json() or {}
    save_data(data)
    return {'status': 'success'}, 200
```

#### Route Access (Direct Python Access):
```python
@app.route('/documents', methods=['GET', 'POST'])
def documents():
    data = load_data()  # Access from session
    
    # Use data to process documents
    for filename in doc_list:
        # Generate documents with `data`
```

---

## 3. DATA FLOW DIAGRAM

```
┌─────────────────────────────────────────────────────────────┐
│                    USER BROWSER                              │
├─────────────────────────────────────────────────────────────┤
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Form Inputs on index.html                           │   │
│  │  (acc_name, acc_age, perm_address, etc.)             │   │
│  └─────────────────┬──────────────────────────────────┬─┘   │
│                    │                                  │       │
│            saveToLocalStorage()        updatePreview()│       │
│                    │                                  │       │
│                    ▼                                  ▼       │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  localStorage['case_data']                           │   │
│  │  {acc_name: "John", acc_age: "30", ...}              │   │
│  └──────────┬───────────────────────────┬──────────────┘   │
│             │                           │                   │
│   Network   │                           │ Direct JS Access  │
│    POST     │                           │ (loadFromLocal    │
│             ▼                           │  Storage())       │
└─────────────────────────────────────────┼───────────────────┘
              │                           │
              │                           │
              ▼                           ▼
┌─────────────────────────────────────────────────────────────┐
│                      SERVER (Flask)                          │
├─────────────────────────────────────────────────────────────┤
│  /api/save_data (POST)                                      │
│       ▼                                                      │
│  session['case_data'] = {acc_name: "John", ...}             │
│       ▲                                                      │
│       │ Access from any route                               │
│  /documents, /generate, /preview, etc.                      │
│                                                              │
│  Access in routes:                                          │
│  data = load_data()  # Returns session['case_data']         │
└─────────────────────────────────────────────────────────────┘
```

---

## 4. PRACTICAL EXAMPLES

### **Example 1: Access Data on documents.html**
```html
<!-- In any HTML file -->
<script>
    function displayAccusedName() {
        const accusedName = localStorage.getItem('case_data');
        if (accusedName) {
            const data = JSON.parse(accusedName);
            document.getElementById('accused-display').innerText = data.acc_name;
        }
    }
    window.addEventListener('load', displayAccusedName);
</script>
```

### **Example 2: Access Data in Python Route**
```python
@app.route('/my_custom_route')
def my_route():
    data = load_data()  # Get from session
    
    accused_name = data.get('acc_name', 'Unknown')
    accused_age = data.get('acc_age', 'N/A')
    
    return render_template('my_template.html', 
                         accused_name=accused_name,
                         accused_age=accused_age)
```

### **Example 3: Generate Document with Stored Fields**
```python
from docx import Document

def generate_arrest_memo():
    data = load_data()
    
    # Load template
    doc = Document('word_templates/Arrest_memo_281.docx')
    
    # Replace placeholders with stored data
    for field in REQUIRED_FIELDS:
        placeholder = f"[{field}]"
        value = data.get(field, "")
        
        for paragraph in doc.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)
    
    return doc
```

---

## 5. AVAILABLE FIELDS

**Accused Details:**
```
acc_name, acc_father, acc_surname, acc_alias, acc_gender, acc_age, 
acc_dob, acc_religion, acc_caste, acc_subcaste, acc_nationality, acc_marital
```

**Occupation:**
```
occ_type, occ_place, occ_income
```

**Addresses:**
```
perm_house, perm_area, perm_village, perm_district, perm_taluka, 
perm_state, perm_pin, curr_address, curr_city, curr_district, 
curr_taluka, curr_state, curr_pin
```

**Contact & ID:**
```
mobile_1, mobile_2, id_type, id_number
```

**Physical Description:**
```
phy_height, phy_build, phy_complexion, phy_eyes, phy_hair, phy_facial_hair,
mark_1, mark_2, old_wounds, other_id_marks
```

**Case Details:**
```
case_* fields, offence_*, bail_*, arrest_*, court_*, etc.
```

---

## 6. KEY POINTS TO REMEMBER

| Aspect | Location | Access Method | Persistence |
|--------|----------|----------------|-------------|
| **Frontend Storage** | Browser localStorage | `localStorage.getItem('case_data')` | Until manually cleared |
| **Backend Storage** | Flask session dict | `load_data()` function | Session lifetime |
| **Sync Method** | Automatic via fetch | POST to `/api/save_data` | Real-time |
| **Cross-Page Access** | Any HTML file | Direct localStorage call | Within same domain |
| **Data Format** | JSON string | Parse with `JSON.parse()` | Full structure preserved |

---

## 7. HOW TO USE IN YOUR FILES

### **To Access in Any HTML File:**
```javascript
// Get all data
let allData = JSON.parse(localStorage.getItem('case_data') || '{}');

// Get specific field
let accusedName = allData.acc_name;

// Use in document
document.getElementById('some-element').innerText = accusedName;
```

### **To Access in Python Route:**
```python
@app.route('/custom_page')
def custom_page():
    data = load_data()  # Returns stored session data
    return render_template('template.html', **data)
```

### **To Access in New Document Generation:**
```python
# Inside any function in app.py
data = load_data()
print(data['acc_name'])  # Direct dictionary access
print(data.get('acc_age', 'N/A'))  # Safe access with default
```

---

## Summary

Your system is **smart and efficient**:
- ✅ **localStorage** keeps user's data persistent in browser
- ✅ **Session** maintains data on server between page loads
- ✅ **Auto-sync** ensures both stay in sync via API calls
- ✅ **Easy access** from any file using simple functions
- ✅ **Scalable** to add new fields without code changes
