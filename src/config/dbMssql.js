const Sequelize = require('sequelize')


const sequelize = new Sequelize(process.env.DB_BASE_MSSQL, process.env.DB_USERNAME_MSSQL, process.env.DB_PASSWORD_MSSQL, {
    host: process.env.DB_HOST_MSSQL,
    dialect: 'mssql',
    logging: false,
    dialectOptions: {
        options: {
            trustedconnection: false,
            encrypt: false,
            trustServerCertificate: false,
            enableArithAbort: false,
        }
    }
});
sequelize.authenticate().then(() => {
    console.log('Connection has been established successfully MSSQL.');
}).catch((error) => {
    console.log('Unable to connect to the database MSSQL: ', error);
});

module.exports = sequelize;