# 🔧 COMPREHENSIVE FIXES - IMPLEMENTATION GUIDE

## Overview of 6 Fixes
1. ✅ Emergency Phone number display in View Registration
2. ✅ Foreigner status visibility in View Registration  
3. ✅ Payment ID save/update functionality
4. ✅ Emergency number printing in Ticket PDF
5. ✅ Foreigner badge in Ticket PDF
6. ✅ Foreigner badge in Reports

---

## FIX 1: Code.gs - Add updatePaymentIdAndEmergency Function

**Location:** Add this new function after `updateTrekkerDetails()` function (around line 828)

**Purpose:** Enable saving Payment ID and Emergency Contact updates from the View Registration form

```javascript
// ====== Admin: Update Payment ID and Emergency Contact ======
function updatePaymentIdAndEmergency(email, trekDate, paymentId, emergencyContact) {
  try {
    email = (email || '').toString().trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };
    const trekDateDMY = formatDateDDMMYYYY(trekDate);
    if (!trekDateDMY) return { success: false, message: 'Invalid trek date.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Registrations');
    if (!sheet) return { success: false, message: 'Registrations sheet not found.' };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'No registration data.' };
    const headers = data[0].map(h => (h || '').toString());

    const emailIdx = headers.findIndex(h => (h || '').toLowerCase() === 'email');
    const trekDateIdx = headers.findIndex(h => (h || '').toLowerCase() === 'trek date' || (h || '').toLowerCase() === 'trekdate');
    const paymentIdIdx = headers.findIndex(h => (h || '').toLowerCase() === 'payment id' || (h || '').toLowerCase() === 'paymentid');
    const emergencyIdx = headers.findIndex(h => (h || '').toLowerCase() === 'emergency contact' || (h || '').toLowerCase() === 'emergency_contact' || (h || '').toLowerCase() === 'emergency');

    if (emailIdx === -1 || trekDateIdx === -1) {
      return { success: false, message: 'Required columns not found.' };
    }

    let updated = false;
    for (let i = 1; i < data.length; i++) {
      if ((data[i][emailIdx] + '').toLowerCase() === email &&
          formatDateDDMMYYYY(data[i][trekDateIdx]) === trekDateDMY) {
        
        if (paymentIdIdx !== -1 && paymentId !== undefined && paymentId !== null) {
          sheet.getRange(i + 1, paymentIdIdx + 1).setValue(paymentId.toString());
          updated = true;
        }
        
        if (emergencyIdx !== -1 && emergencyContact !== undefined && emergencyContact !== null) {
          sheet.getRange(i + 1, emergencyIdx + 1).setValue(emergencyContact.toString());
          updated = true;
        }
      }
    }

    if (!updated) return { success: false, message: 'No matching registration found.' };
    return { success: true, message: 'Payment ID and Emergency Contact updated successfully.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
```

---

## FIX 2: Code.gs - Update getGroupDetails to Include Emergency Contact

**Location:** Modify `getGroupDetails()` function (around line 705)

**Change:** Add emergency contact retrieval to the function

The function ALREADY retrieves it. Verify this section exists (around line 713):
```javascript
EmergencyContact: emergencyIdx !== -1 ? (firstRow[emergencyIdx] || '') : ''
```

✅ This is already in your code.

---

## FIX 3: index.html - Update renderGroupDetails Function

**Location:** Find `function renderGroupDetails(groupData)` (around line 1746)

**What to change:** Add Emergency Contact and Foreigner Status display with Edit buttons

Replace the trekkers table section with this enhanced version:

```javascript
// ====== Render Group Details (UPDATED with Foreigner & Emergency) ======
function renderGroupDetails(groupData) {
  if (!groupData) return;
  
  // Payment ID section - make editable
  const paymentIdHtml = `
    <div class="col-md-6 mb-3">
      <div class="small text-muted">Payment ID</div>
      <div class="d-flex gap-2 align-items-center">
        <input type="text" id="edit-payment-id" class="form-control form-control-custom" 
               value="${groupData['Payment ID'] || ''}" placeholder="Enter Payment ID">
        <button class="btn btn-sm btn-primary-custom" onclick="savePaymentIdAndEmergency()">
          <i class="bi bi-save"></i> Save
        </button>
      </div>
    </div>
  `;

  // Emergency Contact section - make editable
  const emergencyHtml = `
    <div class="col-md-6 mb-3">
      <div class="small text-muted">Emergency Contact</div>
      <div class="d-flex gap-2 align-items-center">
        <input type="tel" id="edit-emergency-contact" class="form-control form-control-custom" 
               value="${groupData.EmergencyContact || ''}" placeholder="Enter Emergency Phone">
        <button class="btn btn-sm btn-primary-custom" onclick="savePaymentIdAndEmergency()">
          <i class="bi bi-save"></i> Save
        </button>
      </div>
    </div>
  `;

  // Existing group info HTML with additions
  document.getElementById('group-info').innerHTML = `
    <div class="row g-3 mb-4">
      <div class="col-md-6">
        <div class="small text-muted">Group Email</div>
        <div class="info-text">${groupData.Email || 'N/A'}</div>
      </div>
      <div class="col-md-6">
        <div class="small text-muted">Trek Date</div>
        <div class="info-text">${groupData['Trek Date'] || 'N/A'}</div>
      </div>
    </div>
    
    <div class="row g-3 mb-4">
      <div class="col-md-6">
        <div class="small text-muted">Amount Remitted</div>
        <div class="info-text">₹ ${groupData['Amount Remitted'] || '0'}</div>
      </div>
      <div class="col-md-6">
        <div class="small text-muted">Ticket Number</div>
        <div class="info-text">${groupData['Ticket No'] || 'Not Issued'}</div>
      </div>
    </div>

    <div class="row g-3 mb-4">
      ${paymentIdHtml}
      ${emergencyHtml}
    </div>

    <div class="row g-3 mb-4">
      <div class="col-md-6">
        <div class="small text-muted">Status</div>
        <div class="info-text">
          <span class="badge ${groupData.Status === 'Ticket Issued' ? 'bg-success' : 'bg-warning'}">
            ${groupData.Status || 'Pending'}
          </span>
        </div>
      </div>
    </div>
  `;

  // Render Trekkers with Foreigner Badge
  const trekkerHtml = groupData.Trekkers.map((trekker, idx) => `
    <tr>
      <td>${idx + 1}</td>
      <td>
        <div class="d-flex align-items-center gap-2">
          <span>${trekker.Name || 'N/A'}</span>
          ${trekker.isForeigner ? '<span class="badge bg-warning text-dark"><i class="bi bi-globe"></i> Foreigner</span>' : ''}
        </div>
      </td>
      <td>${trekker.Gender || 'N/A'}</td>
      <td>${trekker.Age || 'N/A'}</td>
      <td>${trekker['Photo ID Type'] || 'N/A'}</td>
      <td>${trekker['ID Number'] || 'N/A'}</td>
      <td>${trekker['Mobile Number'] || 'N/A'}</td>
      <td>
        <button class="btn btn-sm btn-outline-primary" onclick="showEditModalForTrekker(${idx})">
          <i class="bi bi-pencil"></i> Edit
        </button>
      </td>
    </tr>
  `).join('');

  const trekkerContainer = document.getElementById('trekkers-table-body');
  if (trekkerContainer) {
    trekkerContainer.innerHTML = trekkerHtml;
  }
}

// ====== Save Payment ID and Emergency Contact ======
function savePaymentIdAndEmergency() {
  const paymentId = document.getElementById('edit-payment-id')?.value || '';
  const emergencyContact = document.getElementById('edit-emergency-contact')?.value || '';

  if (!selectedGroupEmail || !selectedTrekDate) {
    showMessage('group-details-message', 'Group email or trek date not set.', 'danger');
    return;
  }

  showMessage('group-details-message', 'Saving...', 'info');

  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.success) {
        showMessage('group-details-message', resp.message || 'Payment ID and Emergency Contact updated!', 'success');
        // Reload group details
        viewGroupDetails(selectedGroupEmail, selectedTrekDate);
      } else {
        showMessage('group-details-message', resp.message || 'Failed to save.', 'danger');
      }
    })
    .withFailureHandler(function() {
      showMessage('group-details-message', 'Failed to save. Please try again.', 'danger');
    })
    .updatePaymentIdAndEmergency(selectedGroupEmail, selectedTrekDate, paymentId, emergencyContact);
}
```

---

## FIX 4: Ticket.html - Add Emergency Number & Foreigner Badge

**Location:** In Ticket.html template

**Add this in the Trekker Details table section:**

```html
<!-- Add to trekker details display -->
<td colspan="2" style="border: 1px solid #ddd; padding: 10px;">
  <strong>Emergency Contact:</strong><br>
  <% if(groupDetails.emergencyContact) { %>
    <%= groupDetails.emergencyContact %>
  <% } else { %>
    Not provided
  <% } %>
</td>

<!-- Update trekker name display to include badge -->
<td style="border: 1px solid #ddd; padding: 10px;">
  <%= trekker.name %>
  <% if(trekker.isForeigner) { %>
    <span style="background: #f59e0b; color: #1f2937; padding: 2px 6px; border-radius: 3px; font-size: 11px; font-weight: bold; margin-left: 5px;">
      🌍 FOREIGNER
    </span>
  <% } %>
</td>
```

---

## FIX 5: TrekkersReport.html - Add Foreigner Badge

**Location:** In TrekkersReport.html template where trekker names are displayed

**Replace trekker name cell with:**

```html
<td style="border: 1px solid #ddd; padding: 10px;">
  <%= trekker.Name %>
  <% if(trekker.isForeigner) { %>
    <span style="background: #f59e0b; color: #1f2937; padding: 2px 6px; border-radius: 3px; font-size: 10px; font-weight: bold; margin-left: 5px; display: inline-block;">
      FOREIGNER
    </span>
  <% } %>
</td>
```

**Add column header:**
```html
<th style="border: 1px solid #ddd; padding: 10px; background: #1e40af; color: white;">Name / Status</th>
```

---

## FIX 6: Code.gs - Update getIssuedTrekkersReport to Include isForeigner

**Location:** In `getIssuedTrekkersReport()` function (around line 1299)

**Verify this code exists:**
```javascript
results.push({
  TicketNo: data[i][ticketNoIdx] || '',
  Name: data[i][nameIdx] || '',
  Gender: data[i][genderIdx] || '',
  Age: data[i][ageIdx] || '',
  IDType: data[i][idTypeIdx] || '',
  IDNumber: toPlainNumberString(data[i][idNumberIdx]) || '',
  Mobile: toPlainNumberString(data[i][mobileIdx]) || '',
  TrekDate: trekDateStr,
  AmountRemitted: amt,
  isForeigner: (isForeignerIdx !== -1 ? ('' + (data[i][isForeignerIdx] || '')).toLowerCase() === 'yes' : false)
});
```

If isForeigner is missing, add it as shown above.

---

## 📋 Testing Checklist

After implementing all fixes:

- [ ] **Fix 1:** Emergency contact displays in View Registration form
- [ ] **Fix 2:** Can edit Emergency contact and Payment ID
- [ ] **Fix 3:** Changes are saved to database
- [ ] **Fix 4:** Foreigner status shows as badge in View Registration
- [ ] **Fix 5:** Emergency number prints in Ticket PDF
- [ ] **Fix 6:** Foreigner badge displays in Ticket PDF
- [ ] **Fix 7:** Foreigner badge shows in Reports

---

## 🔍 Summary of Changes

| Fix # | File | Change | Impact |
|-------|------|--------|--------|
| 1 | Code.gs | Add updatePaymentIdAndEmergency() | Enables saving updates |
| 2 | Code.gs | Verify getGroupDetails() | Already included |
| 3 | index.html | Update renderGroupDetails() | Display + Edit UI |
| 4 | Ticket.html | Add Emergency & Foreigner | Print in PDF |
| 5 | TrekkersReport.html | Add Foreigner badge | Show in report |
| 6 | Code.gs | Update getIssuedTrekkersReport() | Include isForeigner data |

---

## 📍 Where to Find Each Section

- **index.html:** Search for `renderGroupDetails` or `Group Details (Admin)`
- **Code.gs:** Search for `getGroupDetails` or `updateTrekkerDetails`
- **Ticket.html:** Search for `trekker.name` or `groupDetails.emergencyContact`
- **TrekkersReport.html:** Search for table header or trekker name display

