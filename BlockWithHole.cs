// BlockWithHole.cs
using System;
using System.Globalization;
using NXOpen;
using NXOpen.Features;
using NXOpen.UF;

public class BlockWithHole
{
    public static Session theSession;
    public static UFSession theUFSession;
    public static Part workPart;

    public static void Main(string[] args)
    {
        try
        {
            theSession = Session.GetSession();
            theUFSession = UFSession.GetUFSession();
            workPart = theSession.Parts.Work;

            CreateBlockWithHole();
        }
        catch (Exception ex)
        {
            // Покажем ошибку в NX если что-то пойдёт не так
            NXOpen.UI.GetUI().NXMessageBox.Show("BlockWithHole", NXOpen.NXMessageBox.DialogType.Error,
                ex.ToString());
        }
    }

    public static void CreateBlockWithHole()
    {
        // Параметры (можно менять)
        double lengthX = 100.0;
        double lengthY = 50.0;
        double lengthZ = 20.0;
        double holeDiameter = 20.0;

        double centerX = lengthX / 2.0;
        double centerY = lengthY / 2.0;
        double topZ = lengthZ;

        // ===== 1) Создаём блок =====
        BlockFeatureBuilder blockBuilder = workPart.Features.CreateBlockFeatureBuilder(null);
        blockBuilder.Type = BlockFeatureBuilder.Types.OriginAndEdgeLengths;

        // SetOriginAndLengths принимает строки — используем invariant culture чтобы всегда была точка, а не запятая
        blockBuilder.SetOriginAndLengths(
            new Point3d(0.0, 0.0, 0.0),
            lengthX.ToString(CultureInfo.InvariantCulture),
            lengthY.ToString(CultureInfo.InvariantCulture),
            lengthZ.ToString(CultureInfo.InvariantCulture)
        );

        Feature blockFeature = blockBuilder.CommitFeature();
        blockBuilder.Destroy();

        Body body = blockFeature.GetBodies()[0];

        // ===== 2) Находим верхнюю и нижнюю плоские грани через UF AskFaceData =====
        Face topFace = null;
        Face bottomFace = null;

        foreach (Face face in body.GetFaces())
        {
            // Подготовим контейнеры для вывода
            int faceType = 0;                    // тип поверхности
            double[] axisPoint = new double[3];  // точка на оси (или точка на плоскости)
            double[] axisVector = new double[3]; // осевой вектор / нормаль
            double[] bbox = new double[6];
            double r1 = 0.0;
            double r2 = 0.0;
            int flip = 0;

            // AskFaceData заполняет axisVector — для плоскости это нормаль (см. документацию).
            theUFSession.Modl.AskFaceData(face.Tag, out faceType, axisPoint, axisVector, bbox, out r1, out r2, out flip);

            // В UF: код для плоскости = 22 (в разных версиях может быть другое значение; в большинстве справочников плоскость возвращается как planar)
            // Практически в журналах часто проверяют faceType == 1 (в некоторых версиях). Надёжнее — смотреть компонент нормали по Z.
            // Поэтому просто проверим компонент Z нормали.
            if (Math.Abs(axisVector[0]) < 1e-6 && Math.Abs(axisVector[1]) < 1e-6)
            {
                if (axisVector[2] > 0.0)
                    topFace = face;
                else if (axisVector[2] < 0.0)
                    bottomFace = face;
            }

            if (topFace != null && bottomFace != null)
                break;
        }

        if (topFace == null)
            throw new Exception("Верхняя грань не найдена.");
        if (bottomFace == null)
            throw new Exception("Нижняя грань не найдена.");

        // ===== 3) Создаём простое отверстие (HoleFeatureBuilder) =====
        // Используем SetSimpleHole + SetThruFace для сквозного отверстия.
        HoleFeatureBuilder holeBuilder = workPart.Features.CreateHoleFeatureBuilder(null);

        Point3d holeLocation = new Point3d(centerX, centerY, topZ);

        // SetSimpleHole(referencePoint, reverseDirection, placementFace, diameter)
        // placementFace ожидает NXOpen.ISurface, но Face можно привести к ISurface.
        holeBuilder.SetSimpleHole(holeLocation, false, (NXOpen.ISurface)topFace, holeDiameter.ToString(CultureInfo.InvariantCulture));

        // Указываем, что отверстие проходит до нижней грани (сквозное)
        holeBuilder.SetThruFace((NXOpen.ISurface)bottomFace);

        Feature holeFeature = holeBuilder.CommitFeature();
        holeBuilder.Destroy();

        // Обновим представления
        workPart.Views.Refresh();
    }

    // NX Open требует этот метод для правильной выгрузки библиотеки, если вы собираете DLL
    public static int GetUnloadOption(string arg)
    {
        return (int)Session.LibraryUnloadOption.Immediately;
    }
}
