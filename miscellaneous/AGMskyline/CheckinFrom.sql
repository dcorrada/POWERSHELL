/*
Questa query viene lanciata sul DB recuperato da un dump di backup di Snipe
Dovrebbe restituire il numero di checkin/checkout per utente (aka quanti PC ha 
consegnato o sostituito)
*/
SELECT action_logs.created_at AS 'timestamp', action_logs.action_type,
       CONCAT(users.first_name, ' ', users.last_name) AS 'fullname', users.email, users.deleted_at,
       assets.name AS 'asset', assets.serial, status_labels.name AS 'status', assets.notes
FROM action_logs
INNER JOIN users ON action_logs.target_id = users.id
INNER JOIN assets ON action_logs.item_id = assets.id
LEFT JOIN status_labels ON assets.status_id = status_labels.id
WHERE action_logs.action_type IN ('checkin from', 'checkout') AND action_logs.item_type LIKE '%Asset'
-- LIMIT 20
