USE [master]
GO
/****** Object:  Database [payroll_db]    Script Date: 3/10/2016 9:41:50 AM ******/
CREATE DATABASE [payroll_db]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'payroll_db', FILENAME = N'c:\Program Files\Microsoft SQL Server\MSSQL11.SQLEXPRESS\MSSQL\DATA\payroll_db.mdf' , SIZE = 5120KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'payroll_db_log', FILENAME = N'c:\Program Files\Microsoft SQL Server\MSSQL11.SQLEXPRESS\MSSQL\DATA\payroll_db_log.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO
ALTER DATABASE [payroll_db] SET COMPATIBILITY_LEVEL = 110
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [payroll_db].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [payroll_db] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [payroll_db] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [payroll_db] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [payroll_db] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [payroll_db] SET ARITHABORT OFF 
GO
ALTER DATABASE [payroll_db] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [payroll_db] SET AUTO_CREATE_STATISTICS ON 
GO
ALTER DATABASE [payroll_db] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [payroll_db] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [payroll_db] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [payroll_db] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [payroll_db] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [payroll_db] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [payroll_db] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [payroll_db] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [payroll_db] SET  DISABLE_BROKER 
GO
ALTER DATABASE [payroll_db] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [payroll_db] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [payroll_db] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [payroll_db] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [payroll_db] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [payroll_db] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [payroll_db] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [payroll_db] SET RECOVERY SIMPLE 
GO
ALTER DATABASE [payroll_db] SET  MULTI_USER 
GO
ALTER DATABASE [payroll_db] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [payroll_db] SET DB_CHAINING OFF 
GO
ALTER DATABASE [payroll_db] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [payroll_db] SET TARGET_RECOVERY_TIME = 0 SECONDS 
GO
USE [payroll_db]
GO
/****** Object:  Table [dbo].[tbl_AttPayroll]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_AttPayroll](
	[att_id] [int] IDENTITY(1,1) NOT NULL,
	[emp_id] [nvarchar](25) NOT NULL,
	[time_in] [nvarchar](15) NOT NULL,
	[time_out] [nvarchar](15) NOT NULL,
	[reg_hr] [nvarchar](4) NOT NULL,
	[OT] [nvarchar](4) NOT NULL,
	[ND] [nvarchar](4) NOT NULL,
	[att_date] [nvarchar](15) NOT NULL,
	[status] [nvarchar](10) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_department]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_department](
	[dept_id] [int] IDENTITY(1,1) NOT NULL,
	[dept_name] [nvarchar](50) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_docs]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_docs](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[emp_id] [nchar](50) NULL,
	[category] [nchar](50) NULL,
	[description] [nchar](50) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_jobpos]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_jobpos](
	[pos_id] [int] IDENTITY(1,1) NOT NULL,
	[job_pos] [nchar](25) NOT NULL,
	[basic_pay] [nvarchar](25) NOT NULL,
	[department] [nchar](50) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_logindetails]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_logindetails](
	[acc_id] [int] IDENTITY(1,1) NOT NULL,
	[emp_id] [nchar](25) NOT NULL,
	[user_name] [nchar](25) NOT NULL,
	[pass_word] [nchar](25) NOT NULL,
	[access_lvl] [nchar](10) NOT NULL,
	[acc_locked] [nchar](50) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_payroll]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_payroll](
	[salary_id] [int] IDENTITY(1,1) NOT NULL,
	[emp_id] [nvarchar](15) NOT NULL,
	[emp_name] [nvarchar](50) NOT NULL,
	[position] [nvarchar](20) NOT NULL,
	[department] [nvarchar](20) NOT NULL,
	[daterange] [nvarchar](50) NOT NULL,
	[civil_stat] [nvarchar](20) NOT NULL,
	[dependents] [nvarchar](4) NOT NULL,
	[rate_day] [nvarchar](10) NOT NULL,
	[rate_hour] [nvarchar](10) NOT NULL,
	[NH] [nvarchar](10) NOT NULL,
	[OT] [nvarchar](10) NOT NULL,
	[ND] [nvarchar](10) NOT NULL,
	[Grosspay] [nvarchar](20) NOT NULL,
	[SSS] [nvarchar](20) NOT NULL,
	[Philhealth] [nvarchar](20) NOT NULL,
	[Pagibig] [nvarchar](20) NOT NULL,
	[TAX] [nvarchar](20) NOT NULL,
	[net_pay] [nvarchar](20) NOT NULL,
	[additional_deductions] [nvarchar](max) NULL,
	[status] [nchar](10) NOT NULL,
	[date_computed] [nvarchar](25) NOT NULL,
	[month_computed] [nvarchar](25) NOT NULL,
	[time_computed] [nvarchar](20) NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_PayrollData]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_PayrollData](
	[pay_id] [int] IDENTITY(1,1) NOT NULL,
	[emp_id] [nvarchar](15) NOT NULL,
	[tot_RH] [nvarchar](4) NOT NULL,
	[tot_OT] [nvarchar](4) NOT NULL,
	[tot_ND] [nvarchar](4) NOT NULL,
	[daterange] [nvarchar](70) NOT NULL,
	[drange_dur] [nvarchar](4) NOT NULL,
	[status] [nvarchar](4) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_selUserRep]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_selUserRep](
	[rep_id] [int] IDENTITY(1,1) NOT NULL,
	[emp_id] [nvarchar](25) NOT NULL,
	[fn] [nvarchar](30) NOT NULL,
	[ln] [nvarchar](30) NOT NULL,
	[mi] [nvarchar](5) NOT NULL,
	[gr] [nvarchar](7) NOT NULL,
	[jp] [nvarchar](25) NOT NULL,
	[dp] [nvarchar](25) NOT NULL,
	[dh] [nvarchar](35) NOT NULL,
	[dn] [nchar](5) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tbl_user]    Script Date: 3/10/2016 9:41:51 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_user](
	[user_id] [int] IDENTITY(1,1) NOT NULL,
	[emp_id] [nchar](25) NOT NULL,
	[first_name] [nchar](25) NOT NULL,
	[middle_name] [nchar](25) NOT NULL,
	[last_name] [nchar](25) NOT NULL,
	[gender] [nchar](10) NOT NULL,
	[b_date] [nvarchar](50) NOT NULL,
	[age] [nvarchar](50) NOT NULL,
	[c_no] [nvarchar](50) NOT NULL,
	[email] [nvarchar](50) NOT NULL,
	[address] [nvarchar](max) NOT NULL,
	[job_pos] [nchar](25) NOT NULL,
	[dept] [nchar](25) NOT NULL,
	[date_hired] [nvarchar](50) NOT NULL,
	[marital_status] [nchar](15) NOT NULL,
	[dependents] [nvarchar](5) NOT NULL,
	[SSS] [nvarchar](50) NOT NULL,
	[philhealth] [nvarchar](50) NOT NULL,
	[pagibig] [nvarchar](50) NOT NULL,
	[TAX] [nvarchar](50) NOT NULL,
	[DIN] [int] NULL,
	[hr_pr_day] [nvarchar](50) NULL,
	[basic_pay] [nchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
USE [master]
GO
ALTER DATABASE [payroll_db] SET  READ_WRITE 
GO
