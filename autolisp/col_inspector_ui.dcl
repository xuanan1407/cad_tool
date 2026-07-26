// col_inspector_ui.dcl
// Main UI Dialog for CAD Column Inspector Pro
// Copyright (c) 2026 Tran Xuan An

col_inspector : dialog {
  label = "CAD Column Inspector Pro";
  : boxed_column {
    label = "Main Operations";
    : row {
      : button {
        key = "btn_inspect";
        label = "Manual Inspect (COLINSPECT)";
        width = 30;
        fixed_width = true;
      }
    }
    : row {
      : button {
        key = "btn_scan";
        label = "Auto Scan Area (COLSCAN)";
        width = 30;
        fixed_width = true;
      }
    }
  }
  
  : boxed_column {
    label = "Data Management";
    : row {
      : button {
        key = "btn_clear";
        label = "Clear All Data (COLCLEAR)";
        width = 30;
        fixed_width = true;
      }
    }
    : row {
      : button {
        key = "btn_export";
        label = "Export to Excel";
        width = 30;
        fixed_width = true;
      }
    }
  }
  
  : boxed_column {
    label = "Information";
    : text {
      key = "txt_info";
      label = "Click a button above to perform operation";
      alignment = centered;
    }
  }
  
  : row {
    : button {
      key = "btn_help";
      label = "Help";
      width = 10;
      fixed_width = true;
    }
    : spacer { width = 1; }
    : button {
      key = "accept";
      label = "Close";
      is_default = true;
      width = 10;
      fixed_width = true;
    }
  }
}
