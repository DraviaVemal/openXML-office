namespace OpenXMLOffice.Presentation;

public class CommonTools
{
    #region Public Methods

    public static T[][] TransposeArray<T>(T[][] array)
    {
        int vec1 = array.Length;
        int vec2 = array[0].Length;
        T[][] transposedArray = new T[vec2][];

        for (int i = 0; i < vec2; i++)
        {
            transposedArray[i] = new T[vec1];
            for (int j = 0; j < vec1; j++)
            {
                transposedArray[i][j] = array[j][i];
            }
        }
        return transposedArray;
    }

    #endregion Public Methods
}