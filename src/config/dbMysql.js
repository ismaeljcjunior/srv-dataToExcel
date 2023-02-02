const Sequelize = require("sequelize");

const sequelize = new Sequelize(
  process.env.DB_BASE_MYSQL,
  process.env.DB_USERNAME_MYSQL,
  process.env.DB_PASSWORD_MYSQL,
  {
    host: process.env.DB_HOST_MYSQL,
    dialect: process.env.DB_DIALECT_MYSQL,
    logging: false,
    ialectOptions: {
      options: {
        trustedconnection: false,
        encrypt: false,
        trustServerCertificate: false,
        enableArithAbort: false,
      },
    },
  }
);
sequelize
  .authenticate()
  .then(() => {
    console.log("Connection has been established successfully MYSQL.");
  })
  .catch((error) => {
    console.log("Unable to connect to the database MYSQL: ", error);
  });

module.exports = sequelize;
