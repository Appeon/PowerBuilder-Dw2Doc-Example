namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid
{
    public abstract class EntityDefinition<T>
        where T : EntityDefinition<T>
    {
        private T? _previousEntity;

        public T? PreviousEntity
        {
            get { return _previousEntity; }
            set
            {
                _previousEntity = value;
                CalculateOffset();
            }
        }

        public T? NextEntity { get; set; }
        public int Offset { get; protected set; }
        public int IndexOffset { get; protected set; }
        public int Size { get; set; }
        public int Bound => Offset + Size;

        public bool IsEmpty => Objects.Count == 0 && FloatingObjects.Count == 0;
        public IList<VirtualCell> Objects { get; set; }
        public IList<FloatingVirtualCell> FloatingObjects { get; set; }
        /// <summary>
        /// Whether this entity was inserted to fill empty space
        /// </summary>
        public bool IsFiller { get; set; }
        /// <summary>
        /// Whether this entity was inserted to add anchors for floating controls
        /// when such controls go over the existing entities
        /// </summary>
        public bool IsPadding { get; set; }

        public EntityDefinition()
        {
            Objects = new List<VirtualCell>();
            FloatingObjects = new List<FloatingVirtualCell>();
        }

        public virtual void RemoveCell(VirtualCell cell)
        {
            Objects.Remove(cell);
            FloatingObjects.Remove(FloatingObjects.Where(o => o == cell).Single());
        }

        public void CalculateOffset()
        {
            int offset = 0;
            int indexOffset = 0;
            T? previousEntity = PreviousEntity;

            while (previousEntity is not null)
            {
                offset += previousEntity.Size;
                previousEntity = previousEntity.PreviousEntity;
                ++indexOffset;
            }
            Offset = offset;
            IndexOffset = indexOffset;

        }

        public void RemoveFromChain()
        {
            if (Objects.Any())
            {
                throw new InvalidOperationException("Cannor remove an entity with controls in it");
            }
            if (PreviousEntity is not null)
            {
                PreviousEntity.NextEntity = NextEntity;
            }

            if (NextEntity is not null)
            {
                NextEntity.PreviousEntity = PreviousEntity;
            }

            PreviousEntity = null;
            NextEntity = null;
        }
    }
}
