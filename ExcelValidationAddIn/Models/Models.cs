using System;
using System.Collections.Generic;

namespace ExcelValidationAddIn.Models
{
    public class Session
    {
        public string Token { get; set; }
        public DateTime ExpiresAtUtc { get; set; }
        public string TenantId { get; set; }
        public string TenantName { get; set; }

        public bool IsValid()
        {
            // Treat as expired 60s early so a request never fires with a
            // token that dies in flight.
            return !string.IsNullOrEmpty(Token) &&
                   ExpiresAtUtc - DateTime.UtcNow > TimeSpan.FromSeconds(60);
        }

        public bool HasTenant()
        {
            return !string.IsNullOrEmpty(TenantId);
        }
    }

    public class Tenant
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }

    public class LoginResponse
    {
        public string Token { get; set; }
        public DateTime ExpiresAt { get; set; }
    }

    /// <summary>
    /// One row read from the worksheet, keyed by column header (which matches
    /// a property name on the accounting system's document model), plus the
    /// 1-based worksheet row number so results can be mapped back to cells.
    /// </summary>
    public class DocumentRow
    {
        public int SheetRowIndex { get; set; }
        public Dictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }

    public enum ValidationSeverity
    {
        Error,
        Warning
    }

    public class ValidationError
    {
        public int SheetRowIndex { get; set; }
        public string Column { get; set; }
        public string Message { get; set; }
        public ValidationSeverity Severity { get; set; } = ValidationSeverity.Error;
    }

    public class BatchValidationRequest
    {
        public List<DocumentRow> Rows { get; set; } = new List<DocumentRow>();
    }

    public class BatchValidationResponse
    {
        public List<ValidationError> Errors { get; set; } = new List<ValidationError>();
    }

    public enum ConnectivityStatus
    {
        Idle,
        Sending,
        Ok,
        Offline,
        Error
    }

    public class QueueState
    {
        public ConnectivityStatus Status { get; set; } = ConnectivityStatus.Idle;
        public int PendingRowCount { get; set; }
        public DateTime? LastSuccessAtUtc { get; set; }
        public string LastErrorMessage { get; set; }
    }
}
