#!/bin/sh
cdo sellonlatbox B1_TEMP2_481-600.ieg B1_ammersee_TEMP2_481-600.ieg

cdo info B1_ammersee_TEMP2_481-600.ieg > B1_ammersee_TEMP2_481-600.txt

cdo selyear B1_ammersee_TEMP2_481-600.ieg B1_ammersee_TEMP2_custom.ieg

cdo info B1_ammersee_TEMP2_custom.ieg > B1_ammersee_TEMP2_custom.txt

