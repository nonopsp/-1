﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 串口助手1
{
    class AutoSizeForm
    {
        #region 控件缩放
        double formWidth;//窗体原始宽度
        double formHeight;//窗体原始高度
        double scaleX;//水平缩放比例
        double scaleY;//垂直缩放比例
        Dictionary<string, string> ControlsInfo = new Dictionary<string, string>();//控件中心Left,Top,控件Width,控件Height,控件字体Size

        #endregion

        public void GetAllInitInfo(Control ctrlContainer, object form)
        {
            if (ctrlContainer.Parent == form)//获取窗体的高度和宽度
            {
                formWidth = Convert.ToDouble(ctrlContainer.Width);
                formHeight = Convert.ToDouble(ctrlContainer.Height);
            }
            foreach (Control item in ctrlContainer.Controls)
            {
                if (item.Name.Trim() != "")
                {
                    //添加信息：键值：控件名，内容：据左边距离，距顶部距离，控件宽度，控件高度，控件字体。
                    ControlsInfo.Add(item.Name, (item.Left + item.Width / 2) + "," + (item.Top + item.Height / 2) + "," + item.Width + "," + item.Height + "," + item.Font.Size);
                }
                if ((item as UserControl) == null && item.Controls.Count > 0)
                {
                    GetAllInitInfo(item, form);
                }
            }

        }
        public void ControlsChaneInit(Control ctrlContainer)
        {
            scaleX = (Convert.ToDouble(ctrlContainer.Width) / formWidth);
            scaleY = (Convert.ToDouble(ctrlContainer.Height) / formHeight);
        }
        /// <summary>
        /// 改变控件大小
        /// </summary>
        /// <param name="ctrlContainer"></param>
        public void ControlsChange(Control ctrlContainer)
        {
            if (scaleX != 0 || scaleY != 0)
            {
                double[] pos = new double[5];//pos数组保存当前控件中心Left,Top,控件Width,控件Height,控件字体Size
                foreach (Control item in ctrlContainer.Controls)//遍历控件
                {
                    if (item.Name.Trim() != "")//如果控件名不是空，则执行
                    {
                        if ((item as UserControl) == null && item.Controls.Count > 0)//如果不是自定义控件
                        {
                            ControlsChange(item);//循环执行
                        }
                        string[] strs = ControlsInfo[item.Name].Split(',');//从字典中查出的数据，以‘，’分割成字符串组

                        for (int i = 0; i < 5; i++)
                        {
                            pos[i] = Convert.ToDouble(strs[i]);//添加到临时数组
                        }
                        double itemWidth = pos[2] * scaleX;     //计算控件宽度，double类型
                        double itemHeight = pos[3] * scaleY;    //计算控件高度
                        item.Left = Convert.ToInt32(pos[0] * scaleX - itemWidth / 2);//计算控件距离左边距离
                        item.Top = Convert.ToInt32(pos[1] * scaleY - itemHeight / 2);//计算控件距离顶部距离
                        item.Width = Convert.ToInt32(itemWidth);//控件宽度，int类型
                        item.Height = Convert.ToInt32(itemHeight);//控件高度
                        item.Font = new Font(item.Font.Name, float.Parse((pos[4] * Math.Min(scaleX, scaleY)).ToString()));//字体

                    }
                }
            }
        }
    }
}