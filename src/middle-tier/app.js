if (process.env.NODE_ENV !== "production") {
  require("dotenv").config();
}
import * as createError from "http-errors";
import * as path from "path";
import * as cookieParser from "cookie-parser";
import * as logger from "morgan";
import express from "express";
import https from "https";
import { getHttpsServerOptions } from "office-addin-dev-certs";

const app = express();
const port = process.env.API_PORT || "3000";

app.set("port", port);

// view engine setup
/*
app.set("views", path.join(__dirname, "views"));
app.set("view engine", "pug");
*/
app.use(logger("dev"));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());

// Serve static files
app.use(express.static(path.join(process.cwd(), "dist")));

// Serve the client-side task pane files
app.get("/taskpane.html", async (req, res) => {
  return res.sendFile("taskpane.html");
});

app.get("/msgcompose.html", async (req, res) => {
  return res.sendFile("msgcompose.html");
});

app.get("/msgread.html", async (req, res) => {
  return res.sendFile("msgread.html");
});

// Catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(createError(404));
});

// Error handler
app.use(function (err, req, res) {
  res.locals.message = err.message;
  res.locals.error = req.app.get("env") === "development" ? err : {};
  res.status(err.status || 500);
  res.render("error");
});

// Create HTTPS server
getHttpsServerOptions().then((options) => {
  https
    .createServer(options, app)
    .listen(port, () => console.log(`Server running on ${port} in ${process.env.NODE_ENV} mode`));
});
