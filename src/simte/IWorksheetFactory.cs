﻿using System;
using System.Drawing;
using System.IO;
using simte.Common;
using simte.RichText;
using simte.Table;

namespace simte
{
    public interface IWorksheetFactory : IAttachable<IExcelPackage>
    {
        ITableBuilder Table(TableOptions options);
        IRichTextBuilder RichText(Position pos);
        IWorksheetFactory Text(string text, Position pos, Action<ColumnOptionsBuilder> action = null, double? rowHeight = null);

        /// <summary>
        /// Add Image from stream
        /// </summary>
        /// <param name="name"></param>
        /// <param name="pos"></param>
        /// <param name="stream"></param>
        /// <returns>Last cell</returns>
        Position Picture(string name, Position pos, Stream stream);

        Position Picture(string name, Position pos, Image image);

        /// <summary>
        /// Adds an image from the start cell to the end cell by resizing the images
        /// </summary>
        /// <param name="name"></param>
        /// <param name="beg"></param>
        /// <param name="end"></param>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <param name="image"></param>
        /// <returns></returns>
        Position Picture(string name, Position from, Position to, Image image);

        /// <summary>
        /// Indicates where the last action was
        /// </summary>
        int LastRow { get; }


        Size GetPixelsSizeOfCell(Position pos);
        Size GetPixelsSizeOfCellRange(Position from, Position to);
    }
}