using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;
using NPOI.XSSF.UserModel;
using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

[TesterForAttribute(typeof(DwPictureAttributes), typeof(PictureTester))]
public class PictureTester : AbstractAttributeTester<DwPictureAttributes>
{
    protected override AttributeTestResultCollection TestCell(DwPictureAttributes attr, ExportedCell cell)
    {
        throw new NotImplementedException();
    }

    protected override AttributeTestResultCollection TestFloating(DwPictureAttributes attr, ExportedFloatingCell cell)
    {
        var testResults = TestFloatingBase(attr, cell);

        testResults.Add(new(
            cell.Cell.Object.Name,
            "output shape",
            NonNullString,
            cell.OutputShape is null ? NullString : NonNullString));


        while (cell.OutputShape is not null)
        {
            var picture = cell.OutputShape as XSSFPicture;

            testResults.Add(new(
                cell.Cell.Object.Name,
                "shape is picture",
                bool.TrueString,
                (picture is not null).ToString()));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "picture file exists",
                bool.TrueString,
                (File.Exists(attr.FileName) is bool fileExists && fileExists).ToString()));

            if (!fileExists)
                break;

            if (picture is null)
                break;

            var image = Image.FromFile(attr.FileName!);
            using var stream = new MemoryStream();
            var buffer = stream.GetBuffer();

            image.Save(stream, image.RawFormat);

            using var sha1 = System.Security.Cryptography.SHA512.Create();

            testResults.Add(new(
                cell.Cell.Object.Name,
                "picture hashes",
                Convert.ToHexString(sha1.ComputeHash(buffer, 0, buffer.Length)),
                Convert.ToHexString(sha1.ComputeHash(picture.PictureData.Data, 0, picture.PictureData.Data.Length))));

            break;
        }

        return new AttributeTestResultCollection(testResults);
    }
}
