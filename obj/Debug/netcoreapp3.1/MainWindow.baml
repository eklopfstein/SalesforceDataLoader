   M S B A M L   `   `   `  ����E  ADataLoadUI, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:DataLoadUI
DataLoadUI  5   	       DataLoadUI.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   *localclr-namespace:DataLoadUI     5   	     � Title$  
MainWindow��5   	   $	��450��6   $	��760��6)   $	��760��65   $	��450��6D   .��5
      � ���	#FF5F9FD3.+�5   
   �� -   5         ��Name$ lblError��5      $�� ��6"   $��Center=�6-   $	��Top=�6J   ���	#FFF10000 J� RenderTransformOrigin$ 0.564,0.651&�6z   $��	0,325,0,0q�6�   $	��740��6�   $��99��6�   $y�Hidden=�6�   $��Center=�6�   �� 5   
   -   5        $ lblTitle��5      $!��Salesforce Test Data Loader��6"   $��Center=�6H   $��0,10,0,0q�6e   $	��Top=�6w   
�������6�   � 5   
   -   5         �Name$ grdInput��5      $��213,36,218,181q�6!   .+�5      � -   5        $ 	grdLabels��5      $��Center=�6&   $	��193��6C   $	��Top=�6P   $	��120��6h   .+�5      �� -   5        $ lblExcel��5      $��
Excel File��6*   $
��Left=�6?   $	��Top=�6Z   
�������$��	-75,8,0,0q�6�   �� 5      -   5        $ lblUser��5      $��Username��6)   $
��Left=�6<   $	��Top=�6W   
�������$��
-75,39,0,0q�6�   �� 5      -   5        $ lblPassword��5      $��Password��6-   $
��Left=�6@   $	��Top=�6[   
�������$��
-75,73,0,0q�6�   �� 5      -   5        $ lblToken��5      $��Security Token��6*   $
��Left=�6C   $	��Top=�6^   
�������$��
-89,99,0,0q�6�   5      � 5      -	   5        $ 	grdButton��5      $��Center=�6&   $	��193��6C   $��	250,0,0,0q�6P   $	��Top=�6c   $��95��6{   .+�5      �� -
   5         ��Name$ btnFileSearch��5      $��	Find File��60   $	��Top=�6D   $ 0.496,0.609&�6\   $��
23,12,24,0q�6�   $��Center=�6�   5      �� 5      -   5         ��Name$ tbxExcelFile��5      $��Center=�6,   $��0,14,0,0q�6I    �� TextWrapping$
 Wrap=�6[   $	��Top=�6o   $	��120��6�   $ �� HorizontalScrollBarVisibility$ Visible=�6�   $��18��6�   $��18��6�   ${�False��6�   �� 5      -   5        $ tbxUsername��5      $��Center=�6+   $��0,45,0,0q�6H   $
 Wrap=�6Z   $	��Top=�6n   $	��120��6�   $��18��6�   $��18��6�   G� 5      -   5         G�Name$ pbxPassword��5      $��Center=�6/   $��0,77,0,0q�6L   $	��Top=�6^   $	��120��6v   $��18��6�   $��18��6�   6�   G� 5      -   5        $ pbxToken��5      $��Center=�6,   $	��Top=�6I   $	��120��6a   $��	0,105,0,0q�6m   	 G� SelectionOpacity$	 2Q�6�   $��18��6�   $��18��6�   6�   	� 5      -   5        
 	�Name$
 rbnCreateUsers��5      $��Create users��62   $��Center=�6I   $��	0,130,0,0q�6f   $	��Top=�6y   
������� 	� 	GroupName$ btnUser��6�   	� 5      -   5        $
 rbnQueryUsers��5      $��Query users��61   $��Center=�6G   $��	0,150,0,0q�6d   $	��Top=�6w   
�������$ btnUser��6�    d� 	IsChecked$
 TrueY�6�   �� 5      -   5        $ 	btnUpload��5      $��Upload data��6(   $��Center=�6>   $��	0,173,0,0q�6[   $	��Top=�6n   6�   5      �� 5   
   $��Disabled=�6   $
��Auto=�6@   $��
3,258,0,-5q�6c   .��5       �� -   5         ��Name$ tbxLog��5       $��Center=�6(   $��0,3,0,0q�6E   $,��&Run data will appear here after upload��6V   $
��Wrap=�6�   $	��Top=�6�   ${�False��6�   $��Visible=�6�   
�������$
|�True��6
  $
��True��6  $
}�True��6@  5!      5$      5%      