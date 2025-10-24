# Class Crawler - Usage Guide

## What's New

The Class Crawler now displays classification results in a **formatted table** with interactive hover tooltips instead of a long markdown list.

## How to Use

### 1. Run the Tool
- Click the **Class Crawler** button in the IFC Classification pulldown
- This opens the element type selection dialog

### 2. Select Element Types
- Choose one or more element types from the active view
- Click **"Classify Selected"**
- Wait while the AI classifies each type

### 3. Review Results Table
A formatted table will appear with these columns:

| Column | Description |
|--------|-------------|
| **Category** | Revit category (e.g., Walls, Doors) |
| **Family** | Family name |
| **Type** | Type name |
| **Manufacturer** | Manufacturer parameter value |
| **IFC Class** | AI-suggested IFC class (hover to see reasoning) |
| **Predefined Type** | AI-suggested predefined type |
| **Status** | Shows "✓ Ready" when ready to apply |

### 4. View AI Reasoning
- **Hover your mouse** over any **IFC Class** value
- A tooltip will appear showing the AI's reasoning
- This helps you understand why that classification was suggested

### 5. Apply Classifications
- After reviewing the table, a dialog will ask: **"Apply these classifications?"**
- Click **Yes** to apply all classifications to your element types
- Click **No** to cancel without making changes

### 6. Check Results
- A success message shows how many types were updated
- The output window logs the results
- Check your element types to see the applied IFC parameters

## Example

### Before (Old Markdown Display)
```
## 1. Walls - Basic Wall - Generic 200mm
**Original Information:**
- Category: Walls
- Family: Basic Wall
- Type: Generic 200mm
- Manufacturer: 

**Classification Results:**
- IFC Class: IfcWallType
- Predefined Type: SOLIDWALL
- Reasoning: Based on the category "Walls" and the type description...
---
```

### After (New Table Display)
| Category | Family | Type | Manufacturer | IFC Class | Predefined Type | Status |
|----------|--------|------|--------------|-----------|-----------------|--------|
| Walls | Basic Wall | Generic 200mm | - | IfcWallType [hover] | SOLIDWALL | ✓ Ready |

*Hover over "IfcWallType" to see: "Based on the category 'Walls' and the type description..."*

## Tips

- **Hover to Learn:** Always hover over IFC Classes to understand the AI's reasoning
- **Review Before Applying:** The table format makes it easy to spot any incorrect classifications
- **Status Column:** The green checkmark (✓ Ready) indicates the classification is ready to apply
- **Larger Views:** For many results, maximize the output window for better viewing

## Troubleshooting

**Q: Tooltips don't appear?**  
A: Make sure you're hovering directly over the IFC Class text (not empty space)

**Q: Table looks cramped?**  
A: Resize or maximize the output window

**Q: Some reasoning says "No reasoning provided"?**  
A: These won't show tooltips - the AI didn't provide reasoning for that classification

**Q: Can I apply only some classifications?**  
A: Currently, it's all-or-nothing. Review the table and say "No" if you don't want to apply them all.

## Future Enhancements

In future versions, we may add:
- Individual SET buttons for each row (when PyRevit supports it)
- Filtering and sorting capabilities
- Export to Excel
- Confidence scores

## Need Help?

Contact: pyByggstyrning team


