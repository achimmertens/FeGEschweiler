const dhive = require('@hiveio/dhive');

// Erstellen Sie einen neuen Client
let client = new dhive.Client('https://api.hive.blog');

// Definieren Sie den Namen der Community
const communityName = 'hive-10005';

// Abrufen der Posts aus der Community
client.database.getDiscussions('trending', {tag: communityName, limit: 10})
  .then(posts => {
    console.log(posts);
  })
  .catch(error => {
    console.error(error);
  });
