using Microsoft.VisualStudio.TestTools.UnitTesting;
using ArrayLab;

namespace ArrayLab.Tests
{
    [TestClass]
    public class ArrayProcessorTests
    {
        [TestMethod]
        public void ProcessVariant9_NormalArray_ZerosCorrectElements()
        {
            double[] input = { 5.0, 1.2, -10.0, 4.4, 20.0, 0.0 };
            double[] expected = { 5.0, 1.2, -10.0, 0.0, 20.0, 0.0 };
            double[] actual = ArrayProcessor.ProcessVariant9(input);
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ProcessVariant9_AdjacentMinMax_NoChanges()
        {
            double[] input = { 1.0, -5.0, 10.0, 4.0 };
            double[] expected = { 1.0, -5.0, 10.0, 4.0 };
            double[] actual = ArrayProcessor.ProcessVariant9(input);
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ProcessVariant9_ReversedMinMax_ZerosCorrectElements()
        {
            double[] input = { 100.0, 50.0, 20.0, -100.0 };
            double[] expected = { 100.0, 0.0, 0.0, -100.0 };
            double[] actual = ArrayProcessor.ProcessVariant9(input);
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ProcessVariant9_ArrayTooSmall_ReturnsOriginalArray()
        {
            double[] input = { 5.0, 10.0 };
            double[] expected = { 5.0, 10.0 };
            double[] actual = ArrayProcessor.ProcessVariant9(input);
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ProcessVariant9_MultipleMinMax_ZerosMaximumRange()
        {
            double[] input = { 5.0, 2.0, 5.0, 8.0, 4.0, 2.0, 8.0, 5.0 };
            double[] expected = { 5.0, 2.0, 0.0, 0.0, 0.0, 0.0, 8.0, 5.0 };

            double[] actual = ArrayProcessor.ProcessVariant9(input);
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ProcessVariant9_AllElementsEqual_ReturnsOriginalArray()
        {
            double[] input = { 7.0, 7.0, 7.0, 7.0 };
            double[] expected = { 7.0, 7.0, 7.0, 7.0 };

            double[] actual = ArrayProcessor.ProcessVariant9(input);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ProcessVariant9_NullOrEmptyArray_HandlesGracefully()
        {
            double[] emptyInput = new double[0];
            double[] actualEmpty = ArrayProcessor.ProcessVariant9(emptyInput);
            CollectionAssert.AreEqual(new double[0], actualEmpty);

            double[] actualNull = ArrayProcessor.ProcessVariant9(null);
            CollectionAssert.AreEqual(new double[0], actualNull);
        }
    }
}