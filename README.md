[build]
  publish = "."
  functions = "netlify/functions"

[[redirects]]
  from = "/api/storage"
  to = "/.netlify/functions/storage"
  status = 200
  force = true
