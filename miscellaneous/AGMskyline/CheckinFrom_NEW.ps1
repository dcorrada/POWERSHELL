<#
1) Lo script vecchio eseguiva la seguente join da un DB backuppato di SnipeIT, 
un dump che ora l'interfaccia web non riesce più a produrre:

SELECT action_logs.created_at AS 'timestamp', action_logs.action_type,
       CONCAT(users.first_name, ' ', users.last_name) AS 'fullname', users.email, users.deleted_at,
       assets.name AS 'asset', assets.serial, status_labels.name AS 'status', assets.notes
FROM action_logs
INNER JOIN users ON action_logs.target_id = users.id
INNER JOIN assets ON action_logs.item_id = assets.id
LEFT JOIN status_labels ON assets.status_id = status_labels.id
WHERE action_logs.action_type IN ('checkin from', 'checkout') AND action_logs.item_type LIKE '%Asset'


2) Il file csv in output dovrebbe mantenere il seguente formato:

"timestamp","action_type","fullname","email","deleted_at","asset","serial","status","notes"
"2019-11-18 10:57:22","checkout","Gianfranco Di Tommaso","gianfranco.ditommaso@agmsolutions.net","2021-04-02 14:15:41","2FRB6Z2","2FRB6Z2","Assegnato","ex Salvatore Gabrieli"
"2019-11-18 11:32:11","checkin from","Gianfranco Di Tommaso","gianfranco.ditommaso@agmsolutions.net","2021-04-02 14:15:41","2FRB6Z2","2FRB6Z2","Assegnato","ex Salvatore Gabrieli"
"2019-11-18 15:32:37","checkout","Andrea Accolla","andrea.accolla@agmsolutions.net",NULL,"6C9D6Z2-TO","6C9D6Z2","da Assegnare","ex Christian Cammarata"
"2019-11-18 15:36:40","checkout","Gianfranco Di Tommaso","gianfranco.ditommaso@agmsolutions.net","2021-04-02 14:15:41","2FRB6Z2","2FRB6Z2","Assegnato","ex Salvatore Gabrieli"
[...]


3) L'idea sarebbe quella di recuperare le singole tabelle da SnipeIT usando le 
API Restful, come gia' faccio con gli script per le estrazioni asset e utente.
Importo queste tabelle come hash tables e le incrocio simulando le join della 
query SQL.

Per cercare tabelle e sintassi delle API guardare su:
https://snipe-it.readme.io/reference/api-overview
#>
