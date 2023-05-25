const express = require('express');
const fs = require('fs');

const app = express();
app.use(express.json());

// GET /users/:userID
app.get('/users/:userID', (req, res) => {
  const userID = parseInt(req.params.userID);
  fs.readFile('data.json', 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal Server Error');
      return;
    }
    const users = JSON.parse(data).users;
    const user = users.find((user) => user.id === userID);
    if (user) {
      res.json(user);
    } else {
      res.status(404).send('User not found');
    }
  });
});

// GET /posts/:postID
app.get('/posts/:postID', (req, res) => {
  const postID = parseInt(req.params.postID);
  fs.readFile('data.json', 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal Server Error');
      return;
    }
    const posts = JSON.parse(data).posts;
    const post = posts.find((post) => post.id === postID);
    if (post) {
      res.json(post);
    } else {
      res.status(404).send('Post not found');
    }
  });
});

// GET /posts?from=:fromDate&to=:toDate
app.get('/posts', (req, res) => {
  const fromDate = new Date(req.query.from);
  const toDate = new Date(req.query.to);
  if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
    res.status(400).send('Invalid date format');
    return;
  }
  fs.readFile('data.json', 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal Server Error');
      return;
    }
    const posts = JSON.parse(data).posts;
    const filteredPosts = posts.filter((post) => {
      const postDate = new Date(post.last_update);
      return postDate >= fromDate && postDate <= toDate;
    });
    res.json(filteredPosts);
  });
});

// POST /users/:userID/email
app.post('/users/:userID/email', (req, res) => {
  const userID = parseInt(req.params.userID);
  const newEmail = req.body.email;
  fs.readFile('data.json', 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal Server Error');
      return;
    }
    const jsonData = JSON.parse(data);
    const users = jsonData.users;
    const user = users.find((user) => user.id === userID);
    if (user) {
      user.email = newEmail;
      fs.writeFile('data.json', JSON.stringify(jsonData), (err) => {
        if (err) {
          console.error(err);
          res.status(500).send('Internal Server Error');
          return;
        }
        res.send('Email updated successfully');
      });
    } else {
      res.status(404).send('User not found');
    }
  });
});

// PUT /users/:userID/posts
app.put('/users/:userID/posts', (req, res) => {
  const userID = parseInt(req.params.userID);
  const title = req.body.title;
  const body = req.body.body;
  const lastUpdate = new Date().toISOString();
  fs.readFile('data.json', 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal Server Error');
      return;
    }
    const jsonData = JSON.parse(data);
    const users = jsonData.users;
    const posts = jsonData.posts;
    const user = users.find((user) => user.id === userID);
    if (user) {
      const newPost = {
        id: posts.length + 1,
        title: title,
        body: body,
        user_id: userID,
        last_update: lastUpdate
      };
      posts.push(newPost);
      fs.writeFile('data.json', JSON.stringify(jsonData), (err) => {
        if (err) {
          console.error(err);
          res.status(500).send('Internal Server Error');
          return;
        }
        res.send('Post created successfully');
      });
    } else {
      res.status(404).send('User not found');
    }
  });
});

app.listen(3000, () => {
  console.log('Server listening on port 3000');
});
