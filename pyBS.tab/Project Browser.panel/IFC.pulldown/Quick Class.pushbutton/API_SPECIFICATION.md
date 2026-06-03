# IFC Classification API Specification

## Endpoints

### Single Element Classification
```
POST https://n8n.byggstyrning.se/webhook/classify
```

### Batch Element Classification  
```
POST https://n8n.byggstyrning.se/webhook/classify/batch
```

## Request Format

### Headers
```
Content-Type: application/json
```

### Single Element Request (IfcClassifyRequest)
```json
{
  "category": "Walls",
  "family": "Basic Wall", 
  "type": "Generic - 200mm",
  "manufacturer": "Autodesk",
  "description": ""
}
```

### Batch Request (IfcClassifyBatchRequest)
```json
{
  "elements": [
    {
      "category": "Walls",
      "family": "Basic Wall", 
      "type": "Generic - 200mm",
      "manufacturer": "Autodesk",
      "description": ""
    },
    {
      "category": "Doors",
      "family": "Single-Flush",
      "type": "0915 x 2134mm",
      "manufacturer": "Generic",
      "description": ""
    }
  ]
}
```

### Request Fields (IfcClassifyRequest)
- `category`: Revit category name (string)
- `family`: Family name (string)
- `type`: Type name (string) 
- `manufacturer`: Manufacturer parameter value (string, can be empty)
- `description`: Additional description (string, optional)

## Response Format

### Single Element Response (IfcClassifyResponse)
```json
{
  "result": {
    "ifc_class": "IfcWall",
    "predefined_type": "SOLIDWALL",
    "confidence": 0.95,
    "element_id": null
  },
  "processing_time_ms": 45.2
}
```

### Batch Response (IfcClassifyBatchResponse)
```json
{
  "results": [
    {
      "ifc_class": "IfcWall",
      "predefined_type": "SOLIDWALL", 
      "confidence": 0.95,
      "element_id": null
    },
    {
      "ifc_class": "IfcDoor",
      "predefined_type": "DOOR",
      "confidence": 0.87,
      "element_id": null
    }
  ],
  "processing_time_ms": 67.8,
  "total_elements": 2
}
```

### Response Fields

#### IfcClassificationResult
- `ifc_class`: Predicted IFC class (string)
- `predefined_type`: IFC predefined type if applicable (string, optional)
- `confidence`: Confidence score 0.0-1.0 (float)
- `element_id`: Element identifier (string, optional)

#### IfcClassifyResponse
- `result`: Single IfcClassificationResult object
- `processing_time_ms`: Processing time in milliseconds (float)

#### IfcClassifyBatchResponse  
- `results`: Array of IfcClassificationResult objects (same order as input)
- `processing_time_ms`: Total processing time in milliseconds (float)
- `total_elements`: Number of elements processed (int)

### Error Response
```json
{
  "error": "Model not loaded",
  "status": "error",
  "model_loaded": false
}
```

## Expected IFC Classes

The model should predict standard IFC classes such as:
- `IfcWall`
- `IfcSlab` 
- `IfcColumn`
- `IfcDoor`
- `IfcWindow`
- `IfcBeam`
- `IfcRoof`
- `IfcStair`
- `IfcRailing`
- `IfcCovering`
- `IfcAirTerminal`
- `IfcDuctSegment`
- `IfcDuctFitting`
- `IfcPipeSegment` 
- `IfcPipeFitting`
- `IfcSanitaryTerminal`
- `IfcLightFixture`
- `IfcElectricAppliance`
- `IfcUnitaryEquipment`
- `IfcFurniture`
- `IfcBuildingElementProxy` (fallback)

## Implementation Notes

### Performance Expectations
- Support batch processing of 1-100+ elements
- Response time < 5 seconds for typical batches
- Confidence scores should be meaningful (0.0-1.0)

### Fallback Behavior
- If CatBoost model fails, API should return rule-based predictions
- Use confidence < 0.5 to indicate fallback predictions
- Set method field to indicate classification approach used

## Example Payload

Use this for testing the endpoint:

```json
{
  "test": true,
  "elements": [
    {
      "category": "Walls",
      "family": "Basic Wall",
      "type_name": "Generic - 200mm", 
      "manufacturer": "Test"
    }
  ]
}
```

Expected response should classify this as `IfcWall` with high confidence. 