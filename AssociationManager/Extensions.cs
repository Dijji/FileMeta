// Sourced from this blog post http://social.msdn.microsoft.com/forums/en-US/wpf/thread/5909dbcc-9a9f-4260-bc36-de4aa9bbd383/ and copyright of its author
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FileMetadataAssociationManager
{
    public static class Extensions
    {
        public static int FindIndex<T>(this ObservableCollection<T> collection, Predicate<T> predicate)
        {
            int index = 0;
            foreach (var item in collection)
            {
                if (predicate(item))
                    return index;
                index++;
            }
            return -1;
        }

        public static IEnumerable<T> Union<T>(this IEnumerable<T> collection, IEnumerable<T> source, Func<T, T, bool> comparison)
        {
            IEnumerable<T> union = collection.Union(source, new LambdaComparer<T>(comparison));
            return union;
        }

        public static void Sort<T>(this ObservableCollection<T> collection, Comparison<T> comparison)
        {
            var comparer = new Comparer<T>(comparison);

            List<T> sorted = collection.OrderBy(x => x, comparer).ToList();

            for (int i = 0; i < sorted.Count(); i++)
                collection.Move(collection.IndexOf(sorted[i]), i);
        }

        private class Comparer<T> : IComparer<T>
        {
            private readonly Comparison<T> comparison;

            public Comparer(Comparison<T> comparison)
            {
                this.comparison = comparison;
            }

            #region IComparer<T> Members

            public int Compare(T x, T y)
            {
                return comparison.Invoke(x, y);
            }

            #endregion
        }
    }

    public class LambdaComparer<T> : IEqualityComparer<T>
    {
        private readonly Func<T, T, bool> _lambdaComparer;
        private readonly Func<T, int> _lambdaHash;

        public LambdaComparer(Func<T, T, bool> lambdaComparer) :
            this(lambdaComparer, o => 0)
        {
        }

        public LambdaComparer(Func<T, T, bool> lambdaComparer, Func<T, int> lambdaHash)
        {
            if (lambdaComparer == null)
                throw new ArgumentNullException("lambdaComparer");
            if (lambdaHash == null)
                throw new ArgumentNullException("lambdaHash");

            _lambdaComparer = lambdaComparer;
            _lambdaHash = lambdaHash;
        }

        public bool Equals(T x, T y)
        {
            return _lambdaComparer(x, y);
        }

        public int GetHashCode(T obj)
        {
            return _lambdaHash(obj);
        }
    }
}
