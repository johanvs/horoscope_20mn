
--Pour compter les prédictions identiques

SELECT description, count(*) AS nb
FROM horoscopes
GROUP BY description
ORDER BY  count(*) DESC;


--Pour voir quand a été utilisé pour la dernière fois chacune des prédictions

SELECT MAX(h1.id) AS idSame, h.id AS id, idSame - h.id AS diff, FIRST(h.description) AS description,  FIRST(h.jour) AS jour
FROM horoscopes AS h
INNER JOIN horoscopes AS h1 ON  h1.description = h.description AND h1.id < h.id
GROUP BY h.id
ORDER BY h.id;

