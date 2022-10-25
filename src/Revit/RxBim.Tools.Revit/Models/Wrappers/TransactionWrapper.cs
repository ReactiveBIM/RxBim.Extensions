﻿namespace RxBim.Tools.Revit.Models
{
    using Autodesk.Revit.DB;

    /// <inheritdoc cref="ITransactionWrapper" />
    internal class TransactionWrapper : Wrapper<Transaction>, ITransactionWrapper
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TransactionWrapper"/> class.
        /// </summary>
        /// <param name="transaction"><see cref="Transaction"/> instance.</param>
        public TransactionWrapper(Transaction transaction)
            : base(transaction)
        {
        }

        /// <inheritdoc />
        public void Dispose()
        {
            Object.Dispose();
        }

        /// <inheritdoc />
        public void Start()
        {
            if (Object.GetStatus() != TransactionStatus.Started)
                Object.Start();
        }

        /// <inheritdoc />
        public void RollBack()
        {
            if (!IsRolledBack())
                Object.RollBack();
        }

        /// <inheritdoc />
        public bool IsRolledBack()
        {
            return Object.GetStatus() == TransactionStatus.RolledBack;
        }

        /// <inheritdoc />
        public void Commit()
        {
            if (!IsRolledBack() || Object.GetStatus() != TransactionStatus.Committed)
                Object.Commit();
        }
    }
}