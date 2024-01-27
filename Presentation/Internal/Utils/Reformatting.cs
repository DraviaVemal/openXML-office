// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Presentation;

/// <summary>
/// Common Tools useful across presentation library
/// </summary>
public class CommonTools
{


    /// <summary>
    /// Transpose a 2D array
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    /// <param name="array">
    /// </param>
    /// <returns>
    /// </returns>
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


}