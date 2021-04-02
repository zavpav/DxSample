using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using DyDocTestSS.Bl.DataGetters;
using JetBrains.Annotations;

namespace DyDocTestSS.Bl
{

    /// <summary> Привязка колонки к данным </summary>
    public class ColumnDataBind
    {
        public bool IsGroup { get; set; }

        /// <summary> Внутренняя инфомрация по колонкам </summary>
        private List<ColumnDataBindInfoGetter> _columns;

        public ColumnDataBind()
        {
            this.IsGroup = false;
            this._columns = new List<ColumnDataBindInfoGetter>();
        }

        public ColumnDataBind([NotNull] ColumnDataBindInfoGetter getter) : this()
        {
            this._columns.Add(getter);
        }

        public ColumnDataBind([NotNull] [ItemCanBeNull] IEnumerable<ColumnDataBindInfoGetter> getters) : this()
        {
            this._columns.AddRange(getters);
            
            if (this._columns.Count > 1)
                this.IsGroup = true;
        }

        public void AddColumn([NotNull] ColumnDataBindInfoGetter getter)
        {
            this._columns.Add(getter);

            if (this._columns.Count > 1)
                this.IsGroup = true;
        }

        /// <summary> Информация по получению данных для колонки </summary>
        /// <remarks>Может быть одна или несколько колонок по одному описанию</remarks>
        public IEnumerable<ColumnDataBindInfoGetter> Columns {  get { return this._columns; } }
    }

    /// <summary> Абстрактрая привязка получения данных для одной колонки</summary>
    public abstract class ColumnDataBindInfoGetter
    {
    }
    
    /// <summary> Пустая привязка. Ничего делать не надо. </summary>
    public class ColumnDataBindInfoGetterEmpty : ColumnDataBindInfoGetter { }

    /// <summary> Отдельная привязка с ручным вводом данных </summary>
    public class ColumnDataBindInfoGetterUserEdit : ColumnDataBindInfoGetter { }

    /// <summary> Привязка АКСИОКа </summary>
    public class ColumnDataBindInfoGetterAksiok : ColumnDataBindInfoGetter
    {
        /// <summary> Номер свода </summary>
        [CanBeNull]
        public string Svod { get; set; }

        /// <summary> Номер параметра </summary>
        [CanBeNull]
        public string  Param { get; set; }

        [CanBeNull]
        public string RzPrz { get; set; }

        [CanBeNull]
        public string Csr { get; set; }

        [CanBeNull]
        public string Vr { get; set; }

        [CanBeNull]
        public string Kosgu { get; set; }

        [CanBeNull]
        public string Nr { get; set; }

        [CanBeNull]
        public string Sfr { get; set; }

        /// <summary> Код ДК </summary>
        [CanBeNull]
        public string Dk { get; set; }

        /// <summary> Год </summary>
        [CanBeNull]
        public string Year { get; set; }

        /// <summary> Сформировать строку получения данных АКСИОКа в "нашем формате" </summary>
        public string ToStringBind()
        {
            var str = "АКСИОК\n";
            if (this.Svod != null)
                str += "СВОД:" + this.Svod + ";";
            if (this.Param != null)
                str += "ПАРАМЕТР:" + this.Param + ";";
            if (this.RzPrz != null)
                str += "РЗПРЗ:" + this.RzPrz + ";";
            if (this.Csr != null)
                str += "ЦСР:" + this.Csr + ";";
            if (this.Vr != null)
                str += "ВР:" + this.Vr + ";";
            if (this.Kosgu != null)
                str += "КОСГУ:" + this.Kosgu + ";";
            if (this.Nr != null)
                str += "НР:" + this.Nr + ";";
            if (this.Sfr != null)
                str += "СФР:" + this.Sfr + ";";
            if (this.Dk != null)
                str += "Dk:" + this.Dk+ ";";
            return str;
        }

        public ColumnDataBindInfoGetterAksiok Clone()
        {
            return (ColumnDataBindInfoGetterAksiok) this.MemberwiseClone();
        }
    }

    /// <summary> Привязка к данным БР </summary>
    public class ColumnDataBindInfoGetterBr : ColumnDataBindInfoGetter
    {
        /// <summary> Тип БР </summary>
        public enum EnumPrjBrType
        {
            /// <summary> Финансисты </summary>
            Fin,

            /// <summary> Плановики </summary>
            Bp
        }

        public EnumPrjBrType BrType { get; set; }

        /// <summary> Код, с которого получаем данные </summary>
        public string FullSprKey { get; set; }
    }

    /// <summary> "Простая Формула" </summary>
    public class ColumnDataBindInfoGetterSomeFormula : ColumnDataBindInfoGetter
    {
        /// <summary> Объект находиться в процессе построения </summary>
        private bool GetterIsBuilding { get; set; }

        
        public ColumnDataBindInfoGetterSomeFormula()
        {
            this.GetterIsBuilding = true;
            this.CalculatorParameter = Expression.Parameter(typeof(IDataCalculatorInternal), "calc");
            this.PbsCodeParameter = Expression.Parameter(typeof(string), "pbsCode");
        }

        /// <summary> Результирующая функция </summary>
        [CanBeNull]
        public Func<IDataCalculatorInternal, string, decimal> FinalFunction { get; set; }
        
        /// <summary> Параметр калькулятора </summary>
        private ParameterExpression CalculatorParameter { get; set; }

        /// <summary> Параметр Код ПБС </summary>
        private ParameterExpression PbsCodeParameter { get; set; }

        /// <summary> Создать expression получения данных operand </summary>
        private Expression CreateOperand([NotNull] ColumnDataBindInfoGetter operand)
        {
            var operandNew = Expression.New(operand.GetType());
            var binds = new List<MemberBinding>();
            foreach (var property in operand.GetType().GetProperties(BindingFlags.Instance 
                                                                     | BindingFlags.Public 
                                                                     | BindingFlags.SetProperty 
                                                                     | BindingFlags.SetProperty))
            {
                var vl = Expression.Constant(property.GetValue(operand, null), property.PropertyType);
                binds.Add(Expression.Bind(property, vl));
            }

            var createdDataBind = Expression.MemberInit(operandNew, binds.ToArray());


            var funcInfo = typeof(IDataCalculatorInternal).GetMethod("GetDecimal") ?? throw new NotSupportedException("Не нашли метод GetDecimal");

            var expressionPart = Expression.Call(this.CalculatorParameter, funcInfo,
                this.PbsCodeParameter, createdDataBind);
            return expressionPart;
        }

        [CanBeNull]
        private Expression LeftOperand { get; set; }
        
        [CanBeNull]
        private Func<Expression, Expression, BinaryExpression> OperationFunc {get;set; }


        private void CreateOperation([NotNull] ColumnDataBindInfoGetter operand)
        {
            if (this.LeftOperand == null || this.OperationFunc == null)
                throw new NotSupportedException("Ошибка инициализации операции или операнда");

            var r = this.CreateOperand(operand);
            var l = this.LeftOperand;
            // ReSharper disable once PossibleNullReferenceException
            this.LeftOperand = this.OperationFunc(l, r);
        }

        public void MinusOperation([NotNull] ColumnDataBindInfoGetter operand)
        {
            if (!this.GetterIsBuilding)
                throw new NotSupportedException("Ошибка MinusOperation так как объкт уже построен");

            if (this.LeftOperand != null)
                this.CreateOperation(operand);
            else
                this.LeftOperand = this.CreateOperand(operand);

            this.OperationFunc = Expression.Subtract;
        }

        public void PlusOperation([NotNull] ColumnDataBindInfoGetter operand)
        {
            if (!this.GetterIsBuilding)
                throw new NotSupportedException("Ошибка MinusOperation так как объкт уже построен");

            if (this.LeftOperand != null)
                this.CreateOperation(operand);
            else
                this.LeftOperand = this.CreateOperand(operand);

            this.OperationFunc = Expression.Add;
        }

        public void FinishOperand([NotNull] ColumnDataBindInfoGetterAksiok operand)
        {
            if (!this.GetterIsBuilding)
                throw new NotSupportedException("Ошибка Финального оператора так как объкт уже построен");

            this.CreateOperation(operand);
            this.OperationFunc = null;
            this.GetterIsBuilding = false;

            if (this.LeftOperand == null)
                throw new NotSupportedException("Ошибка Финального оператора так как объкт не содержит ничего");
            var final = Expression.Lambda(this.LeftOperand,this.CalculatorParameter, this.PbsCodeParameter);

            this.FinalFunction = (Func<IDataCalculatorInternal, string, decimal>) final.Compile();
        }
    }



}