-- UNDO: Kembalikan CHS1 yang sudah diupdate
BEGIN TRAN;
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/1/16';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/1/24';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/1/27';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/1/37';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/1/38';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/24/10/16';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/24/10/17';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/24/10/24';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/24/10/27';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/24/10/37';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/24/12/38';
UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/3/16';
COMMIT;
