const { readExcel } = require("./readExcel");
const writeExcel = require("./writeExcel");
const express = require("express");
const app = express();
const handlebars = require("express-handlebars");
const http = require("http"); // or 'https' for https:// URLs
const fs = require("fs");
const url = require("url");

var bodyParser = require("body-parser");
var multer = require("multer");

app.use(function (req, res, next) {
  res.header("Access-Control-Allow-Origin", "http://localhost");
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept"
  );
  next();
});
/** Serving from the same express Server
    No cors required */
app.use(express.static("../client"));
app.use(bodyParser.json());
var storage = multer.diskStorage({
  //multers disk storage settings
  destination: function (req, file, cb) {
    cb(null, "./uploads/");
  },
  filename: function (req, file, cb) {
    var datetimestamp = Date.now();
    cb(
      null,
      file.fieldname +
        "-" +
        datetimestamp +
        "." +
        file.originalname.split(".")[file.originalname.split(".").length - 1]
    );
  },
});
var upload = multer({
  //multer settings
  storage: storage,
});
/** API path that will upload the files */

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.engine(
  "hbs",
  handlebars({
    extname: ".hbs",
    defaultLayout: "index.hbs",
    layoutsDir: __dirname + "/views/layouts",
    partialsDir: __dirname + "/views/partials",
  })
);

app.use(express.static("public"));

app.set("views", "./views");
app.set("view engine", "hbs");

app.post("/excel", upload.single("file"), readExcel, function (req, res, next) {
  
});
app.get("/download", function (req, res) {
  const urlParams = new URLSearchParams(req.query.tab);
  let direccion = urlParams.get("direccionExcel");
  const file = `${__dirname}/` + direccion;
  res.download(file); // Set disposition and send it.
});

app.get("/", (req, res) => {

  res.render("./layouts/index", {
    layout: "index",
  });
});

app.listen(8080, () => console.log("Server started on 8080"));
