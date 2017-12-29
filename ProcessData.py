import numpy as np
import pandas as pd
import openpyxl as xl
import treelib


def get_root_list():
    wb_site = xl.load_workbook("Site Information.xlsx")
    ws_site = wb_site['Site Capacity(Hybrid&Packet)']
    df_site = pd.DataFrame(ws_site.values)
    row_count = df_site.shape[0]
    root_list = []
    for index in range(1, row_count):
        site_type = df_site.iloc[index][2]
        if site_type == 'Root':
            root_list.append(df_site.iloc[index, 0])
    return root_list


def find_path(graph, start, end, path=[]):
    path = path + [start]
    if start == end:
        return path
    if start not in graph:
        return None
    for node in graph[start]:
        if node not in path:
            newpath = find_path(graph, node, end, path)
            if newpath != None:
                return newpath
    return None


def generate_topology_tree():
    wb_link = xl.load_workbook("Link Information.xlsx")
    ws_link = wb_link['Link']
    df_link = pd.DataFrame(ws_link.values)
    root_list = get_root_list()
    for root_id in root_list:
        child_list = find_childrens(root_id)


def find_childrens(site_id, df_link):
    pass


if __name__ == '__main__':
    graph = {'A': ['B', 'C'],
             'B': ['D', 'E', 'F'],
             'C': ['H'],
             'E': ['G']}
    print('A->G', find_path(graph, 'A', 'G'))
    print('A->C', find_path(graph, 'A', 'C'))
    print('A->H', find_path(graph, 'A', 'H'))
    print('A->F', find_path(graph, 'A', 'F'))
