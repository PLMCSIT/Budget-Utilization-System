using BUR_UI.Entities;
using System.Collections.Generic;

namespace BUR_UI.Logic
{
    public class Builder
    {
        public List<SAAOModel> FillSAAOModel(List<AccountsModel> acct, List<ABModel> ab)
        {
            List<SAAOModel> SAAO = new List<SAAOModel>();

            for (int i = 0; i < ab.Count; i++)
            {
                SAAO.Insert(i, new SAAOModel()
                {
                    AB = ab[i].ApprovedBudget,
                    Code = ab[i].AccountCode,
                });

                for (int j = 0; j < acct.Count; j++)
                {
                    if (SAAO[i].Code == acct[j].AccountCode)
                        SAAO[i].Amount += acct[j].Amount;
                }
            }

            return SAAO;
        }
        public List<SAAOModel> FillMonthlyModel(List<AccountsModel> acct, List<ABModel> ab)
        {
            List<SAAOModel> Monthly = new List<SAAOModel>();

            for (int i = 0; i < ab.Count; i++)
            {
                Monthly.Insert(i, new SAAOModel()
                {
                    AB = ab[i].ApprovedBudget,
                    Code = ab[i].AccountCode,
                });

                for (int j = 0; j < acct.Count; j++)
                {
                    if (Monthly[i].Code == acct[j].AccountCode)
                        Monthly[i].Amount += acct[j].Amount;
                }
            }

            return Monthly;
        }
    }
}
