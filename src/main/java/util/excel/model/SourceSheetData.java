package util.excel.model;

import lombok.Data;

@Data
public class SourceSheetData {
   private String serviceName;
   private String costCenter;
   private String iccCode;
   private Double cost;
   private String serviceOwner;
}
