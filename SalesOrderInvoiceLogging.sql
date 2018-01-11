
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[SalesOrderInvoiceLogging](
	[TransactionID] [bigint] IDENTITY(1,1) NOT NULL,
	[SalesOrder] [nvarchar](100) NULL,
	[DocumentType] [nvarchar](100) NULL,
	[SentForValidationDate] [datetime] NULL,
	[ItemsProcessedForValidation] [nvarchar](50) NULL,
	[ItemsInvalidForValidation] [nvarchar](50) NULL,
	[ItemsRejectedWithWarningsForValidation] [nvarchar](50) NULL,
	[ItemsProcessedWithWarningsForValidation] [nvarchar](50) NULL,
	[ErrorNumberForValidation] [nvarchar](50) NULL,
	[ErrorDescriptionForValidation] [nvarchar](1000) NULL,
	[SentForValidation] [bit] NULL,
	[SentForInvoicing] [bit] NULL,
	[SentForInvoicingDate] [datetime] NULL,
	[InvoiceNumber] [nvarchar](100) NULL,
	[TrnYear] [nvarchar](50) NULL,
	[TrnMonth] [nvarchar](50) NULL,
	[Register] [nvarchar](50) NULL,
	[WarningNumber] [nvarchar](50) NULL,
	[WarningDescription] [nvarchar](50) NULL,
	[Processed] [nvarchar](50) NULL,
	[GlYear] [nvarchar](50) NULL,
	[GlPeriod] [nvarchar](50) NULL,
	[GlJournal] [nvarchar](50) NULL,
	[ItemsProcessedForInvoicing] [nvarchar](50) NULL,
	[ItemsInvalidForInvoicing] [nvarchar](50) NULL,
	[ItemsRejectedWithWarningsForInvoicing] [nvarchar](50) NULL,
	[ItemsProcessedWithWarningsForInvoicing] [nvarchar](50) NULL,
	[ErrorNumberForInvoicing] [nvarchar](50) NULL,
	[ErrorDescriptionForInvoicing] [nvarchar](1000) NULL
) ON [PRIMARY]

GO


