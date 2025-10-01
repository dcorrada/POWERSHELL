--
-- MySQL server version 15.1 
-- Distribution 10.11.14-MariaDB, for debian-linux-gnu (aarch64)
-- Host: localhost    Database: AGMskyline
--

-- Table structure for table `ADcomputers`
CREATE TABLE `ADcomputers` (
  `ID` varchar(8) NOT NULL,
  `HOSTNAME` varchar(80) DEFAULT NULL,
  `OU` varchar(80) DEFAULT NULL,
  `OS` varchar(80) DEFAULT NULL,
  `LOGONDATE` date DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `HOSTNAME` (`HOSTNAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='computer a dominio';

-- Table structure for table `ADusers`
CREATE TABLE `ADusers` (
  `ID` varchar(8) NOT NULL,
  `USRNAME` varchar(80) NOT NULL,
  `UPN` varchar(256) DEFAULT NULL,
  `FULLNAME` varchar(256) DEFAULT NULL,
  `OU` text,
  `CREATED` date DEFAULT NULL,
  `LASTLOGON` date DEFAULT NULL,
  `DESCRIPTION` text,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `USRNAME` (`USRNAME`),
  KEY `FULLNAME` (`FULLNAME`),
  KEY `UPN` (`UPN`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='utenze a dominio';

-- Table structure for table `AzureDevices`
CREATE TABLE `AzureDevices` (
  `ID` varchar(8) NOT NULL,
  `LOGONDATE` date DEFAULT NULL,
  `OSTYPE` varchar(80) DEFAULT NULL,
  `OSVER` varchar(80) DEFAULT NULL,
  `HOSTNAME` varchar(80) DEFAULT NULL,
  `OWNER` varchar(80) DEFAULT NULL,
  `MAIL` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `HOSTNAME` (`HOSTNAME`),
  KEY `OWNER` (`OWNER`),
  KEY `MAIL` (`MAIL`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='asset registrarti su AzureAD';

-- Table structure for table `CheckinFrom`
CREATE TABLE `CheckinFrom` (
  `ID` varchar(8) NOT NULL,
  `UPTIME` date DEFAULT NULL,
  `CHECKINOUT` varchar(20)DEFAULT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `MAIL` varchar(80) DEFAULT NULL,
  `USR_STATUS` varchar(20) DEFAULT NULL,
  `HOSTNAME` varchar(80) DEFAULT NULL,
  `SERIAL` varchar(80) DEFAULT NULL,
  `ASSET_STATUS` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `FULLNAME` (`FULLNAME`),
  KEY `MAIL` (`MAIL`),
  KEY `HOSTNAME` (`HOSTNAME`),
  KEY `SERIAL` (`SERIAL`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='consegne e assegnazioni PC';

-- Table structure for table `DLmembers`
CREATE TABLE `DLmembers` (
  `ID` varchar(8) NOT NULL,
  `DLNAME` varchar(80) DEFAULT NULL,
  `DLDESC` text,
  `DLMAIL` varchar(256) DEFAULT NULL,
  `DLCREATED` date NOT NULL,
  `DLSECURITY` enum('True','False') NOT NULL,
  `DLMAILENABLED` enum('True','False') NOT NULL,
  `DLTYPE` varchar(80) DEFAULT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `EMAIL` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `DLNAME` (`DLNAME`),
  KEY `EMAIL` (`EMAIL`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='membri delle distribution list';

-- Table structure for table `DymoLabel`
DROP TABLE IF EXISTS `DymoLabel`;
CREATE TABLE `DymoLabel` (
  `ID` varchar(8) NOT NULL,
  `HOSTNAME` varchar(80) NOT NULL,
  `UPTIME` date NOT NULL,
  `QRCODE` varchar(256) NOT NULL,
  PRIMARY KEY (`ID`),
  KEY `HOSTNAME` (`HOSTNAME`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='etichette applicate agli asset';

-- Table structure for table `EstrazioneAsset`
DROP TABLE IF EXISTS `EstrazioneAsset`;
CREATE TABLE `EstrazioneAsset` (
  `ID` varchar(8) NOT NULL,
  `HOSTNAME` varchar(80) DEFAULT NULL,
  `STATUS` varchar(80) DEFAULT NULL,
  `MODEL` varchar(80) DEFAULT NULL,
  `SERIAL` varchar(80) DEFAULT NULL,
  `NUMPROT` varchar(80) DEFAULT NULL,
  `LOCATION` varchar(80) DEFAULT NULL,
  `UPDATED` date DEFAULT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `USRNAME` varchar(80) DEFAULT NULL,
  `CPU` varchar(20) DEFAULT NULL,
  `RAM` varchar(20) DEFAULT NULL,
  `SSD` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `HOSTNAME` (`HOSTNAME`),
  UNIQUE KEY `SERIAL` (`SERIAL`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='estrazione asset da SnipeIT';

-- Table structure for table `EstrazioneUtenti`
DROP TABLE IF EXISTS `EstrazioneUtenti`;
CREATE TABLE `EstrazioneUtenti` (
  `ID` varchar(8) NOT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `USRNAME` varchar(80) DEFAULT NULL,
  `EMAIL` varchar(80) DEFAULT NULL,
  `PHONE` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `EMAIL` (`EMAIL`),
  KEY `FULLNAME` (`FULLNAME`),
  KEY `USRNAME` (`USRNAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='estrazione utenti da SnipeIT';

-- Table structure for table `GFIparsed`
DROP TABLE IF EXISTS `GFIparsed`;
CREATE TABLE `GFIparsed` (
  `ID` varchar(8) NOT NULL,
  `USER` varchar(80) DEFAULT NULL,
  `HOSTNAME` varchar(80) DEFAULT NULL,
  `OS` varchar(80) DEFAULT NULL,
  `ALIAS` varchar(80) DEFAULT NULL,
  `AGENT` varchar(80) DEFAULT NULL,
  `AV` varchar(256) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `USER` (`USER`),
  KEY `HOSTNAME` (`HOSTNAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='estrazione da GFI per antivirus gestiti';

-- Table structure for table `PwdExpire`
DROP TABLE IF EXISTS `PwdExpire`;
CREATE TABLE `PwdExpire` (
  `ID` varchar(8) NOT NULL,
  `USRNAME` varchar(80) NOT NULL,
  `FULLNAME` varchar(256) NOT NULL,
  `ACCOUNT_EXPDATE` date DEFAULT NULL,
  `PWD_LASTSET` date DEFAULT NULL,
  `PWD_EXPIRED` enum('True','False','NeverExpire','') NOT NULL,
  `PWD_EXPDATE` date DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `USRNAME` (`USRNAME`),
  KEY `FULLNAME` (`FULLNAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='scadenza password';

-- Table structure for table `SchedeAssunzione`
DROP TABLE IF EXISTS `SchedeAssunzione`;
CREATE TABLE `SchedeAssunzione` (
  `ID` varchar(8) NOT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `MAIL` varchar(80) DEFAULT NULL,
  `USRNAME` varchar(80) DEFAULT NULL,
  `ROLE` varchar(80) DEFAULT NULL,
  `SCOPE` varchar(80) DEFAULT NULL,
  `STATUS` varchar(80) DEFAULT NULL,
  `LOCATION` varchar(80) DEFAULT NULL,
  `STARTED` date DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `MAIL` (`MAIL`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='estrazione da file excel di schede assunzioni e cessazioni';

-- Table structure for table `SchedeSIM`
DROP TABLE IF EXISTS `SchedeSIM`;
CREATE TABLE `SchedeSIM` (
  `ID` varchar(8) NOT NULL,
  `FULLNAME` varchar(80) NOT NULL,
  `PHONENUM` varchar(80) NOT NULL,
  `ACTIVATION` date NOT NULL,
  `NOTES` text,
  PRIMARY KEY (`ID`),
  KEY `FULLNAME` (`FULLNAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='assegnazioni SIM gestite da Max';

-- Table structure for table `SchedeTelefoni`
DROP TABLE IF EXISTS `SchedeTelefoni`;
CREATE TABLE `SchedeTelefoni` (
  `ID` varchar(8) NOT NULL,
  `FILENAME` varchar(80) NOT NULL,
  `UPDATED` date NOT NULL,
  `STATUS` varchar(80) DEFAULT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `DESCRIPTION` text,
  PRIMARY KEY (`ID`),
  KEY `FILENAME` (`FILENAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='estrazione documenti assegnazione telefoni';

-- Table structure for table `ThirdPartiesLicenses`
DROP TABLE IF EXISTS `ThirdPartiesLicenses`;
CREATE TABLE `ThirdPartiesLicenses` (
  `ID` varchar(8) NOT NULL,
  `PRODUCT` varchar(80) DEFAULT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `MAIL` varchar(80) DEFAULT NULL,
  `EXPIRE` date DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `MAIL` (`MAIL`),
  KEY `FULLNAME` (`FULLNAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='licenze assegnate non Microsoft';

-- Table structure for table `TrendMicroparsed`
DROP TABLE IF EXISTS `TrendMicroparsed`;
CREATE TABLE `TrendMicroparsed` (
  `ID` varchar(8) NOT NULL,
  `USER` varchar(80) NOT NULL,
  `HOSTNAME` varchar(80) NOT NULL,
  `OS` varchar(80) DEFAULT NULL,
  `AGENT` varchar(80) DEFAULT NULL,
  `ENGINE` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `USR` (`USER`),
  KEY `HOST` (`HOSTNAME`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='Trend Micro antivirus';

-- Table structure for table `UpdatedTables`
DROP TABLE IF EXISTS `UpdatedTables`;
CREATE TABLE `UpdatedTables` (
  `ID` int NOT NULL AUTO_INCREMENT,
  `ATABLE` varchar(80)NOT NULL,
  `UPDATED` date NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='aggiornamenti delle tabelle';

-- Table structure for table `Xhosts`
DROP TABLE IF EXISTS `Xhosts`;
CREATE TABLE `Xhosts` (
  `ID` int NOT NULL AUTO_INCREMENT,
  `HOSTNAME` varchar(80) DEFAULT NULL,
  `ESTRAZIONEASSET` varchar(8) NOT NULL,
  `ADCOMPUTERS` varchar(8),
  `AZUREDEVICES` varchar(8),
  `GFIPARSED` varchar(8),
  `TRENDMICROPARSED` varchar(8),
  `CHECKINFROM` varchar(8),
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Table structure for table `Xusers`
DROP TABLE IF EXISTS `Xusers`;
CREATE TABLE `Xusers` (
  `ID` int NOT NULL AUTO_INCREMENT,
  `FULLNAME` varchar(256) DEFAULT NULL,
  `O365LICENSES` varchar(8),
  `AZUREDEVICES` varchar(8),
  `DLMEMBERS` varchar(8),
  `SCHEDEASSUNZIONE` varchar(8),
  `SCHEDETELEFONI` varchar(8),
  `SCHEDESIM` varchar(8),
  `THIRDPARTIESLICENSES` varchar(8),
  `ESTRAZIONEUTENTI` varchar(8),
  `CHECKINFROM` varchar(8),
  `ADUSERS` varchar(8) NOT NULL,
  `PWDEXPIRE` varchar(8),
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Table structure for table `o365licenses`
DROP TABLE IF EXISTS `o365licenses`;
CREATE TABLE `o365licenses` (
  `ID` varchar(8) NOT NULL,
  `FULLNAME` varchar(80) DEFAULT NULL,
  `MAIL` varchar(80) DEFAULT NULL,
  `STARTED` date DEFAULT NULL,
  `LICENSE` text,
  PRIMARY KEY (`ID`),
  KEY `MAIL` (`MAIL`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='licenze assegnate Microsoft ';
