const dhive = require('@hiveio/dhive');

// Erstellen Sie einen neuen Client
let client = new dhive.Client('https://api.hive.blog');

// Definieren Sie den Namen der Community
const communityName = 'hive-10005';

// Abrufen der Posts aus der Community
client.database.getDiscussions('trending', { tag: communityName, limit: 10 })
    .then(posts => {
        console.log(posts);
        posts.forEach(post => {
            console.log("Author = ", post.author, "Permlink = ", post.permlink);


            client.database.call('get_content_replies', [post.author, post.permlink]).then(comments => {
                comments.forEach(comment => {
                    console.log(`Author: ${comment.author}`);
                    console.log(`Permlink: ${comment.permlink}`);
                    console.log(`Body: ${comment.body}`);
                    console.log('-------------------------');
                });
            }).catch(e => {
                console.error('Fehler beim Abrufen der Kommentare:', e);
            });
            // Abrufen der Kommentare zu jedem Post
            // client.call('condenser_api.get_content_replies', [post.author, post.permlink])
            //   .then(comments => {
            //     console.log('Kommentare zum Post ' + post.permlink + ':', comments);
            //   })
            //   .catch(error => {
            //     console.error('Fehler beim Abrufen der Kommentare:', error);
            //   });
        })

    }).catch(error => {
        console.error('Fehler beim Abrufen der Posts:', error);
    });