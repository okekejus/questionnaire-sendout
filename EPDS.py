
# -*- coding: utf-8 -*-
"""
@author: Justin Okeke
EPDS Evaluation

"""

import obs_data


class Lsq2Epds():
    def __init__(self, obs_ids, api_param, cut_off):
        self.redcap_lsq2 = self._redcap_lsq2_epds(api_param)
        self.redcap_lsq2 = self._lsq2_epds(self.redcap_lsq2)
        self.fu_ids = self._get_lsq2_ids(
            obs_ids, self.redcap_lsq2, cut_off
        )

    def _redcap_lsq2_epds(self, api_param):
        """Download LSQ2 EPDS data from REDCap

        Parameters
        ----------
        api_param : str
            API associated with LSQ

        Returns
        -------
        redcap_lsq2_epds : pandas.dataframe
            Dataframe containing LSQ2 EPDS data

        """
        redcap_lsq2_epds = obs_data.redcap_data(
            api_param,
            {
                'fields[0]': 'obs_study_id',
                'fields[1]': 'lwk_funny',
                'fields[2]': 'lwk_lookfo',
                'fields[3]': 'lwk_blame',
                'fields[4]': 'lwk_anxio',
                'fields[5]': 'lwk_scare',
                'fields[6]': 'lwk_top',
                'fields[7]': 'lwk_sleep',
                'fields[8]': 'lwk_miser',
                'fields[9]': 'lwk_cryin',
                'fields[10]': 'lwk_harm',
            }
        )

        return redcap_lsq2_epds

    def _lsq2_epds(self, dataframe):
        """LSQ2 EPDS Calculations

        Parameters
        ----------
        dataframe : pandas.dataframe
            Dataframe with LSQ2 EPDS data but without final calculations

        Returns
        -------
        dataframe : pandas.dataframe
            Dataframe with EPDS calculations

        """
        dataframe['lwk_funny'].replace(
            {'1': 0, '2': 1, '3': 2, '4': 3}, inplace=True
        )
        dataframe['lwk_lookfo'].replace(
            {'1': 0, '2': 1, '3': 2, '4': 3}, inplace=True
        )
        dataframe['lwk_blame'].replace(
            {'2': 3, '3': 2, '1': 1, '4': 0}, inplace=True
        )
        dataframe['lwk_anxio'].replace(
            {'2': 0, '1': 1, '3': 2, '4': 3}, inplace=True
        )
        dataframe['lwk_scare'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lwk_top'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lwk_sleep'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lwk_miser'].replace(
            {'2': 3, '3': 2, '1': 1, '4': 0}, inplace=True
        )
        dataframe['lwk_cryin'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lwk_harm'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['epds'] = dataframe[
            [
                'lwk_funny', 'lwk_lookfo', 'lwk_blame', 'lwk_anxio',
                'lwk_scare', 'lwk_top', 'lwk_sleep', 'lwk_miser',
                'lwk_cryin', 'lwk_funny'
            ]
        ].sum(axis=1)

        dataframe['lwk_harm_label'] = dataframe['lwk_harm'].replace(
            {
                3: 'Yes, quite often',
                2: 'Sometimes',
                1: 'Hardly ever',
                0: 'Never'
            }
        )

        return dataframe

    def _get_lsq2_ids(self, obs_ids, dataframe, cut_off):
        """LSQ2 subjects to followup based on EPDS

        Get list of LSQ2 subjects that need to be followed up based on their
        EPDS score or answer to the question about harming themselves

        Parameters
        ----------
        obs_ids : list of strings
            OBS IDs to evaluate if they need to be followed up
        dataframe : pandas.dataframe
            LSQ data with EPDS
        cut_off : int
            EPDS cut off where values greater than or equal to will need to
            be followed up

        Returns
        -------
        list of strings
            OBS IDs who need to be followed up based on their EPDS

        """
        id_dataframe = dataframe[dataframe['obs_study_id'].isin(obs_ids)]

        fu_harm = (
            (id_dataframe['lwk_harm'] == 1)
            | (id_dataframe['lwk_harm'] == 2)
            | (id_dataframe['lwk_harm'] == 3)
        )
        fu_cut_off = (id_dataframe['epds'] >= cut_off)

        fu_dataframe = id_dataframe.loc[
            (fu_harm | fu_cut_off), ['obs_study_id', 'epds', 'lwk_harm_label']
        ]

        return fu_dataframe.values.tolist()


class Lsq3Epds():
    def __init__(self, obs_ids, api_param, cut_off):
        self.redcap_lsq3 = self._redcap_lsq3_epds(api_param)
        self.redcap_lsq3 = self._lsq3_epds(self.redcap_lsq3)
        self.fu_ids = self._get_lsq3_ids(
            obs_ids, self.redcap_lsq3, cut_off
        )

    def _redcap_lsq3_epds(self, api_param):
        """Download LSQ3 EPDS data from REDCap

        Parameters
        ----------
        api_param : str
            API associated with LSQ

        Returns
        -------
        redcap_lsq3_epds : pandas.dataframe
            Dataframe containing LSQ3 EPDS data

        """
        redcap_lsq3_epds = obs_data.redcap_data(
            api_param,
            {
                'fields[0]': 'obs_study_id',
                'fields[1]': 'lweek_laugh',
                'fields[2]': 'lweek_enjoy',
                'fields[3]': 'lweek_blame',
                'fields[4]': 'lweek_anxious',
                'fields[5]': 'lweek_panic',
                'fields[6]': 'lweek_top',
                'fields[7]': 'lweek_unhappy',
                'fields[8]': 'lweek_miserable',
                'fields[9]': 'lweek_crying',
                'fields[10]': 'lweek_harming'
            }
        )

        return redcap_lsq3_epds

    def _lsq3_epds(self, dataframe):
        """LSQ3 EPDS Calculations

        Parameters
        ----------
        dataframe : pandas.dataframe
            Dataframe with LSQ3 EPDS data but without final calculations

        Returns
        -------
        dataframe : pandas.dataframe
            Dataframe with EPDS calculations

        """
        dataframe['lweek_laugh'].replace(
            {'1': 0, '2': 1, '3': 2, '4': 3}, inplace=True
        )
        dataframe['lweek_enjoy'].replace(
            {'1': 0, '2': 1, '3': 2, '4': 3}, inplace=True
        )
        dataframe['lweek_blame'].replace(
            {'2': 3, '3': 2, '1': 1, '4': 0}, inplace=True
        )
        dataframe['lweek_anxious'].replace(
            {'1': 0, '2': 1, '3': 2, '4': 3}, inplace=True
        )
        dataframe['lweek_panic'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lweek_top'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lweek_unhappy'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lweek_miserable'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lweek_crying'].replace(
            {'1': 3, '2': 2, '3': 1, '4': 0}, inplace=True
        )
        dataframe['lweek_harming'].replace(
            {'4': 3, '5': 2, '6': 1, '7': 0}, inplace=True
        )

        dataframe['epds'] = dataframe[
            [
                'lweek_laugh', 'lweek_enjoy', 'lweek_blame', 'lweek_anxious',
                'lweek_panic', 'lweek_top', 'lweek_unhappy', 'lweek_miserable',
                'lweek_crying', 'lweek_harming'
            ]
        ].sum(axis=1)

        dataframe['lweek_harming_lablel'] = dataframe['lweek_harming'].replace(
            {
                3: 'Yes, quite often',
                2: 'Sometimes',
                1: 'Hardly ever',
                0: 'Never'
            }
        )

        return dataframe

    def _get_lsq3_ids(self, obs_ids, dataframe, cut_off):
        """LSQ3 subjects to followup based on EPDS

        Get list of LSQ3 subjects that need to be followed up based on their
        EPDS score or answer to the question about harming themselves

        Parameters
        ----------
        obs_ids : list of strings
            OBS IDs to evaluate if they need to be followed up
        dataframe : pandas.dataframe
            LSQ data with EPDS
        cut_off : int
            EPDS cut off where values greater than or equal to will need to
            be followed up

        Returns
        -------
        list of strings
            OBS IDs who need to be followed up based on their EPDS

        """
        id_dataframe = dataframe[dataframe['obs_study_id'].isin(obs_ids)]

        fu_harm = (
            (id_dataframe['lweek_harming'] == 1)
            | (id_dataframe['lweek_harming'] == 2)
            | (id_dataframe['lweek_harming'] == 3)
        )
        fu_cut_off = (id_dataframe['epds'] >= cut_off)

        fu_dataframe = id_dataframe.loc[
            (fu_harm | fu_cut_off),
            ['obs_study_id', 'epds', 'lweek_harming_lablel']
        ]

        return fu_dataframe.values.tolist()
