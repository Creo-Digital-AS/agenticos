-- Hand the agenticos-backend DB back to the upstream (vstorm) migration chain.
-- Undoes Creo's 0101_session_rotated_at; upstream's 0102_session_rotated_at re-adds the same column.
BEGIN;

SELECT version_num FROM alembic_version;  -- expect 0101_session_rotated_at

ALTER TABLE sessions DROP COLUMN IF EXISTS rotated_at;

UPDATE alembic_version
   SET version_num = '0100_directory_groups'
 WHERE version_num = '0101_session_rotated_at';

SELECT version_num FROM alembic_version;  -- expect 0100_directory_groups

COMMIT;
