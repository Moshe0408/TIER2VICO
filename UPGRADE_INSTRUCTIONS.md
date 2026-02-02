# שדרוג מדריכים - Drag & Drop ומחיקה

## שינויים נדרשים ב-Dashboard_App.py:

### 1. הוספת פונקציית מחיקה (לפני saveGuide)

```javascript
async function deleteGuide(catId, guideId) {
    if(!confirm('האם למחוק מדריך זה?')) return;
    
    guides_data.forEach(c => {
        if(c.guides) c.guides = c.guides.filter(g => g.id != guideId);
        if(c.subCategories) {
            c.subCategories.forEach(s => {
                if(s.guides) s.guides = s.guides.filter(g => g.id != guideId);
            });
        }
    });
    
    await syncGuides();
    update();
    alert('המדריך נמחק בהצלחה');
}
```

### 2. שינוי renderGuidesList - הוספת כפתור מחיקה

חפש את החלק שבו נוצרת רשימת המדריכים והוסף כפתור מחיקה:

```javascript
// ליד כפתור העריכה, הוסף:
<button onclick="event.stopPropagation(); deleteGuide('${catId}', '${g.id}')" 
        style="padding:8px; background:#ef4444; border:none; border-radius:8px; cursor:pointer; margin-left:8px;">
    🗑️
</button>
```

### 3. הוספת Drag & Drop Zone

הוסף בסוף פונקציית init():

```javascript
// Setup Drag & Drop for guides
const guideSection = document.querySelector('#guides-section'); // או איזה אזור שמציג מדריכים
if(guideSection) {
    guideSection.addEventListener('dragover', (e) => {
        e.preventDefault();
        guideSection.style.background = 'rgba(var(--accent-rgb), 0.1)';
        guideSection.style.border = '2px dashed var(--accent)';
    });
    
    guideSection.addEventListener('dragleave', (e) => {
        guideSection.style.background = '';
        guideSection.style.border = '';
    });
    
    guideSection.addEventListener('drop', async (e) => {
        e.preventDefault();
        guideSection.style.background = '';
        guideSection.style.border = '';
        
        const files = Array.from(e.dataTransfer.files);
        if(files.length === 0) return;
        
        // קבל קטגוריה מהמשתמש
        const catId = prompt('הזן ID קטגוריה למדריכים:') || selectedCatId;
        if(!catId) {
            alert('יש לבחור קטגוריה תחילה');
            return;
        }
        
        for(let file of files) {
            await processFileToGuide(file, catId);
        }
        
        await syncGuides();
        update();
        alert(`${files.length} מדריכים נוספו בהצלחה!`);
    });
}
```

### 4. פונקציית עיבוד קובץ למדריך

```javascript
async function processFileToGuide(file, catId) {
    try {
        // Upload file
        const formData = new FormData();
        formData.append('file', file);
        const uploadResp = await fetch('/api/upload', { method: 'POST', body: formData });
        const uploadData = await uploadResp.json();
        
        // Extract content
        const extractResp = await fetch('/api/extract-content', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url: uploadData.url })
        });
        const extractData = await extractResp.json();
        
        if(!extractData.content) return;
        
        // Extract images from content
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = extractData.content;
        const imgs = tempDiv.querySelectorAll('img');
        const images = Array.from(imgs).map(img => img.getAttribute('src')).filter(Boolean);
        
        // Create guide
        const cat = guides_data.find(c => c.id == catId);
        if(!cat) return;
        
        const guideObj = {
            id: Date.now().toString() + Math.random(),
            title: file.name.replace(/\.(docx?|pdf)$/i, ''),
            content: extractData.content,
            images: images
        };
        
        if(!cat.guides) cat.guides = [];
        cat.guides.push(guideObj);
        
    } catch(e) {
        console.error('Error processing file:', file.name, e);
    }
}
```

## הוראות יישום:

1. פתח את `Dashboard_App.py`
2. מצא את הסקריפט הראשי (אחרי `<script>`)
3. הוסף את הפונקציות לעיל
4. מצא את הקוד שמציג רשימת מדריכים והוסף כפתור מחיקה
5. שמור והעלה ל-GIT

## בדיקה:

- נסה לגרור קובץ DOCX לאזור המדריכים
- נסה למחוק מדריך
- וודא שהכל עובד
