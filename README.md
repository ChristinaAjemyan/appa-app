# Express SQLite Application

A Node.js Express application with SQLite database for managing users and posts.

## Features

- RESTful API endpoints
- SQLite database integration
- CORS support
- Environment configuration
- User management (CRUD operations)
- Post management (CRUD operations)

## Project Structure

```
appa_node/
├── src/
│   ├── database.js       # Database configuration and initialization
│   └── routes/
│       ├── users.js      # User routes
│       └── posts.js      # Post routes
├── server.js             # Main application file
├── package.json          # Project dependencies
├── .env                  # Environment variables
├── .gitignore           # Git ignore file
└── README.md            # This file
```

## Installation

1. Install dependencies:
```bash
npm install
```

2. Start the server:
```bash
npm start
```

For development with auto-reload:
```bash
npm run dev
```

The server will run on `http://localhost:3000`

## API Endpoints

### Users
- `GET /api/users` - Get all users
- `GET /api/users/:id` - Get user by ID
- `POST /api/users` - Create new user
- `PUT /api/users/:id` - Update user
- `DELETE /api/users/:id` - Delete user

### Posts
- `GET /api/posts` - Get all posts
- `GET /api/posts/:id` - Get post by ID
- `POST /api/posts` - Create new post
- `PUT /api/posts/:id` - Update post
- `DELETE /api/posts/:id` - Delete post

### Health Check
- `GET /api/health` - Server status

## Example Usage

### Create a User
```bash
curl -X POST http://localhost:3000/api/users \
  -H "Content-Type: application/json" \
  -d '{"name": "John Doe", "email": "john@example.com"}'
```

### Create a Post
```bash
curl -X POST http://localhost:3000/api/posts \
  -H "Content-Type: application/json" \
  -d '{"user_id": 1, "title": "My First Post", "content": "This is my first post"}'
```

### Get All Users
```bash
curl http://localhost:3000/api/users
```

### Get All Posts
```bash
curl http://localhost:3000/api/posts
```

## Database

The application uses SQLite with two main tables:

- **users**: Stores user information (id, name, email, created_at)
- **posts**: Stores post information (id, user_id, title, content, created_at)

The database file (`database.db`) is automatically created on first run.

## Environment Variables

- `PORT` - Server port (default: 3000)
- `NODE_ENV` - Environment (development/production)

## Dependencies

- **express** - Web framework
- **sqlite3** - SQLite database driver
- **cors** - Cross-Origin Resource Sharing middleware
- **dotenv** - Environment variable management

## Development Dependencies

- **nodemon** - Auto-reload server during development
