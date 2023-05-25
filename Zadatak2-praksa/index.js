const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.json());

// GET /users/:userID/excel
app.get('/users/:userID/excel', (req, res) => {
  const userID = parseInt(req.params.userID);
  
  fs.readFile('data.json', 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal Server Error');
      return;
    }
    const jsonData = JSON.parse(data);
    const user = jsonData.users.find((user) => user.id === userID);
    
    if (user) {
      // Create a new Excel workbook
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('User Details');
      
      // Define cell styles for formatting
      const titleStyle = {
        font: { bold: true, size: 12 },
        alignment: { horizontal: 'center' }
      };
      const valueStyle = {
        font: { size: 12 }
      };
      
      // Add headers for user details
      worksheet.getCell('A1').value = 'Name';
      worksheet.getCell('A1').style = titleStyle;
      worksheet.getCell('B1').value = 'Email';
      worksheet.getCell('B1').style = titleStyle;
      
      // Populate user details
      worksheet.getCell('A2').value = user.name;
      worksheet.getCell('A2').style = valueStyle;
      worksheet.getCell('B2').value = user.email;
      worksheet.getCell('B2').style = valueStyle;
      
      // Add headers for posts
      worksheet.getCell('D1').value = 'Post ID';
      worksheet.getCell('D1').style = titleStyle;
      worksheet.getCell('E1').value = 'Title';
      worksheet.getCell('E1').style = titleStyle;
      worksheet.getCell('F1').value = 'Body';
      worksheet.getCell('F1').style = titleStyle;
      
      // Find and populate posts by the user
      const userPosts = jsonData.posts.filter((post) => post.user_id === userID);
      userPosts.forEach((post, index) => {
        const rowIndex = index + 2; // Start from row 2
        
        worksheet.getCell(`D${rowIndex}`).value = post.id;
        worksheet.getCell(`D${rowIndex}`).style = valueStyle;
        worksheet.getCell(`E${rowIndex}`).value = post.title;
        worksheet.getCell(`E${rowIndex}`).style = valueStyle;
        worksheet.getCell(`F${rowIndex}`).value = post.body;
        worksheet.getCell(`F${rowIndex}`).style = valueStyle;
        
        // Auto-fit column widths based on cell values
        worksheet.getColumn('D').width = 10;
        worksheet.getColumn('E').width = 15;
        worksheet.getColumn('F').width = 30;
      });
      
      // Auto-fit column widths for user details
      worksheet.getColumn('A').width = 15;
      worksheet.getColumn('B').width = 25;
      
      // Generate a unique filename for the Excel file
      const filename = `user_${userID}_details.xlsx`;
      
      // Save the workbook to disk
      workbook.xlsx.writeFile(filename)
        .then(() => {
          console.log(`Excel file '${filename}' generated successfully`);
          res.download(filename, (err) => {
            if (err) {
              console.error(err);
              res.status(500).send('Internal Server Error');
            }
            // Delete the file after sending the response
            fs.unlink(filename, (err) => {
              if (err) {
                console.error(err);
              }
              console.log(`Excel file '${filename}' deleted`);
            });
          });
        })
        .catch((err) => {
          console.error(err);
          res.status(500).send('Internal Server Error');
        });
    } else {
      res.status(404).send('User not found');
    }
  });
});


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
