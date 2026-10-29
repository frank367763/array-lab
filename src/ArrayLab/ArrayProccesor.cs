using System;
using System.Linq;

namespace ArrayLab
{
    public class ArrayProcessor
    {
        public static double[] ProcessVariant9(double[] input)
        {
            if (input == null || input.Length < 3) return input?.ToArray() ?? new double[0];

            double[] result = input.ToArray();
            double minVal = result.Min();
            double maxVal = result.Max();

            if (minVal == maxVal) return result;

            int firstMin = Array.IndexOf(result, minVal);
            int lastMin = Array.LastIndexOf(result, minVal);
            int firstMax = Array.IndexOf(result, maxVal);
            int lastMax = Array.LastIndexOf(result, maxVal);

            int dist1 = Math.Abs(lastMax - firstMin);
            int dist2 = Math.Abs(lastMin - firstMax);

            int leftIdx = dist1 >= dist2 ? Math.Min(firstMin, lastMax) : Math.Min(firstMax, lastMin);
            int rightIdx = dist1 >= dist2 ? Math.Max(firstMin, lastMax) : Math.Max(firstMax, lastMin);

            if (rightIdx - leftIdx > 1)
            {
                for (int i = leftIdx + 1; i < rightIdx; i++) result[i] = 0.0;
            }

            return result;
        }
    }
}
