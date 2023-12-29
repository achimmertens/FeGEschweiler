const dhive = require('@hiveio/dhive');

// Erstellen Sie einen neuen Client
let client = new dhive.Client('https://api.hive.blog');

// Definieren Sie den Autor und den Permlink des Posts
const author = 'advertisingbot';
const permlink = 'test-achim-please-ignore';

// Rufen Sie die Kommentare ab
client.database.call('get_content_replies', [author, permlink]).then(comments => {
  comments.forEach(comment => {
    console.log(`Author: ${comment.author}`);
    console.log(`Permlink: ${comment.permlink}`);
    console.log(`Body: ${comment.body}`);
    console.log('-------------------------');
  });
}).catch(e => {
  console.error('Fehler beim Abrufen der Kommentare:', e);
});
