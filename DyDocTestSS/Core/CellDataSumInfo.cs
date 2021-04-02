using JetBrains.Annotations;

namespace Sphaera.Bp.Bl.Excel
{
    /// <summary> Описание суммы в ячейке </summary>
    public interface ICellDataSumInfo
    {
        /// <summary> Системное описание информации о сумме </summary>
        [NotNull]
        string SysInfo();
    }

    /// <summary> Дополнительное описание суммы в ячейке отчёта. Отображается хинтом при отображении отчёта. </summary>
    public class CellDataSumInfoDetailSumm : ICellDataSumInfo
    {
        /// <summary> Собственно, сумма </summary>
        public readonly decimal? Sm;

        /// <summary> Описание суммы </summary>
        [NotNull, UsedImplicitly]
        public readonly string SmInfo;


        /// <summary> Описание суммы в ячейке. Отображается хинтом при отображении excel. </summary>
        public CellDataSumInfoDetailSumm([NotNull] string smInfo, decimal? sm)
        {
            this.Sm = sm;
            this.SmInfo = smInfo;
        }

        public string SysInfo()
        {
            return string.Format("Простое описание {0} \t {1}", this.SmInfo, this.Sm);
        }
    }

    /// <summary> Пункт меню для вызова мультфильма отображающего детальную информацию по сумме. </summary>
    public abstract class CellDataSumInfoCartoon : ICellDataSumInfo
    {

        /// <summary> Получить информацию о мультфильме </summary>
        [NotNull]
        public abstract ICartoonAlgorithmInfo CartoonAlgorithmInfo();

        public string SysInfo()
        {
            return string.Format("Мультфильм " + this.CartoonAlgorithmInfo().CartoonInfo());
        }
    }

    public interface ICartoonAlgorithmInfo
    {
        object CartoonInfo();
    }
}