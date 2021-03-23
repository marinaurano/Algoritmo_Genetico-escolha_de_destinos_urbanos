using System;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Globalization;


namespace ag_1
{       
    class Program
    {
        /*---------------------------------------------------------
        declaração de variáveis globais 
        usei a notação: inicio com a letra M (matriz) e V (vetor))
        as variáveis globais são utilizadas na rotina principal e
        na função contendo a rotina da função objetivo
        *///---------------------------------------------------------        
        //matriz contendo os dados de entrada
        static double[,] mdados = new double[2197, 87]; /* 2197 = 2196 registros + 1; 87 = 86 colunas de dados no input + 1 */
        // vetor contendo o código final de cada viagem
        static Int32[] vcod = new int[2197];
        //rotina de entrada ou rotina principal        
         static double[,] mpi, ms, mc, mm,mfinal;
        static double fo,fomax,fomaxg,fomin,fominant,foming,pcruz,pmut,alpha,corte;
        static Int32 tp, nite, ite, npop,imin,imax,cont,contr,nh,h1,h2,duracaoh,ncorte;
        static DateTime inicio, atual,duracao;
        static double[] vsol,vsolt,vlmin,vlmax,vsolmax,vsolmaxg,vsolmin,vsolming;
        //static Random aleat = new Random(Guid.NewGuid().GetHashCode());
        static Random aleat = new Random((int)DateTime.Now.Ticks);
        static StringDictionary listasd = new StringDictionary();
        static void Main(string[] args)
        {
            CultureInfo cult = new CultureInfo("pt-BR");
            //string dta = DateTime.Now.ToString("dd/MM HH:mm:ss", cult);
            /*listasd.Add("2.981", "9 0 8 9");
            listasd.Add("2.348392", "9 0 4 30");
            bool teste = false;
            teste=listasd.ContainsKey("2.981");

            //listasd.Add("2.981", "2");*/
            //-----------------------------------------------------------
            #region algoritmos_iniciais
            // DECLARAÇÃO DE VARIÁVEIS DA ROTINA PRINCIPAL-----------
            string[] lines = System.IO.File.ReadAllLines(@"limites.txt");
            string[] lines2 = System.IO.File.ReadAllLines(@"input.txt");
            //quantidade de linhas (dados) no arquivo input
            Int32 ndados = lines2.Length;

            vsolmax = new double[581]; /* 581 = 580 parametros + 1 */
            vsolmaxg = new double[581];
            vsolmin = new double[581];
            vsolming = new double[581];
            vlmin = new double[581];
            vlmax = new double[581];
            vsol = new double[581];
            vsolt = new double[581];
            ncorte = 0;
            double fmin = 10000000;
            fomaxg = -1;
            foming = 1000000;
            string st1;
            //tamano da população,quant de iterações; iteração corrente 
            //tamanho da população...
            contr = 0;
            npop = 100;
            alpha = 0.5;//parâmetro do cruzamento
            pmut = 0.2;
            //matrizes que serão utilizadas no AG                    
            ms = new double[npop + 1, 581];
            mc = new double[npop + 1, 581];
            mm = new double[npop + 1, 581];
            // ENTRADA DE DADOS -----------------------------------------
            // lendo os dados contendo os limites de cada parâmetro
            for (var i = 1; i <= 580; i++)
            {
                st1 = lines[i - 1];
                string[] vst1 = st1.Split('\t');
                vlmin[i] = Convert.ToDouble(vst1[0]);
                vlmax[i] = Convert.ToDouble(vst1[1]);
               // Console.WriteLine("{0} {1}", vlmin[i].ToString(), vlmax[i].ToString());

            }
            
            // lendo a matriz input            
            for (var i = 1; i <= ndados; i++)
            {
                st1 = lines2[i - 1];
                string[] vst1 = st1.Split('#');
                for (var j = 1; j <= 86; j++) { mdados[i, j] = Convert.ToDouble(vst1[j - 1]); } /*  = colunas de dados */
                vcod[i] = Convert.ToInt32(mdados[i, 1]);
                //após passar a primeira coluna para o vetor vcod, coloquei todos os valores
                //igual a 1 para facilitar no algoritmo da função objetivo.
                mdados[i, 1] = 1; 
            }
            Console.WriteLine("Tamanho dos dados:" + ndados);
            Console.WriteLine("Algoritmo Genético aplicado ao problema O/D");
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("Digite o valor de corte da Função Objetivo");
            corte = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("Digite ENTER para iniciar e ESC para encerrar o processo");
            Console.ReadKey();            
          

            #endregion

            //INÍCIO DO AG
            Console.WriteLine("Inicio..");
            Console.WriteLine((DateTime.Now.ToString("dd/MM HH:mm",cult)));
            Console.WriteLine("----------------------------------------------");
           
            populacao_inicial();

            ite = 0;

            int final = 0;
                       
            while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape))
            {
                ite++;
                aleat = new Random((int)DateTime.Now.Ticks);
                selecao();
                cruzamento();
                mutacao();
              

            } // fim do loop principal do AG

            Console.WriteLine("Fim..");
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine((DateTime.Now.ToString("dd/MM HH:mm", cult)));
            vsolming[0]= foming;
            /*using (StreamWriter writer = new StreamWriter(@"melhorsolucao.txt"))
            {
                foreach (var value in vsolming)
                {
                    writer.WriteLine(value);
                }
            }
            Console.WriteLine("A melhor solução encontra-se no arquivo melhorsolucao.txt");
            Console.WriteLine("Digite ENTER para Fechar");
            Console.ReadKey();*/
            Console.WriteLine("Ordenando as 100 melhores soluções...");
            ordenando();
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("Salvando as 100 melhores soluções...");
            salvando();
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("As 100 melhores soluções encontram-se no arquivo melhores_solucoes.xlsx");
            Console.WriteLine("As abas terão os nomes no seguinte formato: \"dia_mes_hora_minuto(valor de corte)\" ");
            Console.WriteLine("Digite ENTER para Fechar");
            Console.ReadKey();
            /*string path = Environment.CurrentDirectory;
            path += "\\melhores_solucoes.xlsx";
            ExcelPackage arquivoexcel = new ExcelPackage(new FileInfo(@path));
            ExcelWorkbook planilha = arquivoexcel.Workbook;            
            ExcelWorksheet aba;
            
            st1 = DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString();
            planilha.Worksheets.Add(st1);
            planilha.Worksheets.MoveToStart(st1);
            aba = planilha.Worksheets[1];
            for (var j = 1; j <= 363; j++) aba.Column(j).Width = 20; /* 363 = 361 parametros (19*19) + 2 *//*
            for (var j = 3; j <= 363; j++) aba.Cells[1,j].Value = "param " + (j - 2).ToString();
            aba.Cells[1, 1].Value = "Soluções";
            aba.Cells[1, 2].Value = "F obj";
            for (var j = 1; j <= 363; j++) aba.Cells[1, j].Style.HorizontalAlignment= OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            int cont = 1;
            foreach (DictionaryEntry solucao in listasd)
            {
                cont++;
                aba.Cells[cont, 1].Value = cont-1;
                solucao.Key.ToString();
                aba.Cells[cont, 2].Value = Convert.ToDouble(solucao.Key.ToString());
                st1 = solucao.Value.ToString();
                string[] vst1 = st1.Split(';');
                for (var j = 0; j <= 360; j++) aba.Cells[cont, j + 3].Value = Convert.ToDouble(vst1[j]); /* 360 = 361 parametros - 1 *//*
            }
            
            arquivoexcel.Save();*/

        }
        static double Fobj(double[] solucao)
        {
            double[,] mp = new double[36, 18]; /* 35 = 34 equacoes completas +1; 18 = 17 parametros especificos +1 */
            double[,] mu = new double[2197, 36]; /* 2197 = 2196 registros + 1; 36 = 35 equacoes + 1 */
            double[] v1 = new double[18]; /* 18 = 17 parametros especificios + 1 */
            double[] v2 = new double[18]; /* 18 = 17 parametros especificos + 1 */
            double[] v3 = new double[36]; /* 36 = 35 equacoes + 1 */
            double[] v4 = new double[36]; /* 36 = 35 equacoes + 1 */
            double[] vprob = new double[2197];
            double sp, somaprob = 0;
            double v5, soma, generico, distancia1;

            /*PARTE NOVA DO CÓGIGO - LEITURA DAS VARIÁVEIS INCLUÍDAS (1) E EXCLUÍDAS (0)*/

                       
            int k = 0;
            for (var i = 2; i <= 35; i++) /*35 = 35 equacoes */
            {
                k = k + 17;
                for (var j = 1; j <= 17; j++)
                {
                    mp[i, j] = solucao[k - 17 + j]; /* 17 = 17 parametros especificos; mp é a matriz com os coeficientes que serão estimados */
                    mp[1, j] = 1;
                }
                generico = solucao[580]; /* Para deixar o 18º parametro cte em todas as equacoes (variavel generica). 324 = 19 * 17 + 1 */
                distancia1 = solucao[579];             
            }
            generico = solucao[580]; /* Para deixar o 18º parametro cte em todas as equacoes (variavel generica). 324 = 19 * 17 + 1 */
            distancia1 = solucao[579];

            for (var i = 1; i <= 2196; i++)
            {
                var j = 1;
                if (j == 1)
                {
                    v5 = mdados[i, (51 + j)];
                    mu[i, j] = generico * v5 + distancia1 * mdados[i, (16 + j)];
                    j++;
                }
                for (j = 2; j <= 35; j++)
                {
                    {
                        for (k = 1; k <= 16; k++)                          /* Cria a matriz v1 com os dados que vão ser utilizados em cada equação, inclusive as colunas que entram em uma das equações (variaveis do destino) */
                        {
                            v1[k] = mdados[i, k];
                        }
                        k = 17;
                        if (k == 17)
                        {
                            v1[k] = mdados[i, 16 + j];
                            k++;
                        }
                        if (k == 18)
                        {
                            v5 = mdados[i, (51 + j)];
                        }
                    }
                    for (k = 1; k <= 17; k++) { v2[k] = mp[j, k]; } /* Cria a matriz v2 com os coeficientes que serão estimados */
                    sp = 0;
                    for (k = 1; k <= 17; k++) { sp = sp + v1[k] * v2[k]; } /* Multiplica a matriz de dados com a matriz de coeficientes.*/
                    if (sp > 100) { sp = 100; }
                    if (sp < -100) { sp = -100; }
                    v5 = mdados[i, (51 + j)];
                    mu[i, j] = sp + generico * v5; /* mu é a matriz de utilidades*/
                }
            }

            somaprob = 0;
            for (var i = 1; i <= 2196; i++)
            {
                for (var j = 1; j <= 35; j++) { v3[j] = mu[i, j]; }
                soma = 0;
                for (var j = 1; j <= 35; j++) { soma = soma + Math.Exp(v3[j]); }
                vprob[i] = Math.Exp(mu[i, vcod[i]]) / soma;
                if (vprob[i] < 0.00000000001) { vprob[i] = 0.00000000001; }
                somaprob = somaprob + Math.Log(vprob[i]);
            }

            return Math.Abs(somaprob);
        }

        static void selecao()
        {
            Int32 aleat1, aleat2;
            
            //caso ocorra de uma convergência local, a matriz mm é resetada e o algoritmo 
            //passa para a linha seguinte "popresetada" por meio do comando goto, verificar 
            //mais abaixo...
            //popresetada: 
            if (ite>1) { fominant = fomin; }
            //avaliando cada solução..
            for (var i = 1; i <= npop; i++)
            {
                for (var j = 1; j <= 580; j++) { vsol[j] = mm[i, j]; }
                fo = Fobj(vsol);                
                mm[i, 0] = fo;
                //armazenando as soluções com valores inferiores ao valor de corte
                if (fo < corte)
                {
                    string stfo = fo.ToString();
                    bool repetido = false;
                    repetido = listasd.ContainsKey(stfo);
                    if (repetido == false)
                    {
                        string st1 = "";
                        for (var j = 1; j <= 580; j++) st1 = st1+mm[i, j].ToString() + ";";
                        listasd.Add(stfo, st1);
                        ncorte++;
                    }
                   
                    
                    
                }
            }
            //pegando melhor e pior solução da população
            atualizafoming:
            fomin = mm[1, 0];
            fomax = mm[1, 0];
            for (var i = 2; i <= npop; i++)
            {
                if (mm[i,0]<fomin)
                {
                    fomin = mm[i, 0];
                    for (var j = 1; j <= 580; j++) { vsolmin[j] = mm[i, j]; }
                    imin = i;
                }
                if (mm[i, 0] > fomax)
                {
                    fomax = mm[i, 0];
                    for (var j = 1; j <= 580; j++) { vsolmax[j] = mm[i, j]; }
                    imax = i;
                }
            }
                      
           
            
            //se aparecer um indivíduo na população atual melhor que todos até então
            //atualiza melhor de todos, caso contrário, joga o melhor de todos 
            //na posição do pior da população atual. Uma espécie de elitismo para um
            //indivíduo.
           
            if (fomin<foming)
            {
                foming = fomin;
                for (var j = 1; j <= 580; j++) { vsolming[j] = mm[imin, j]; }
            }            
            if (fomin>foming)
            {
                for (var j = 1; j <= 580; j++) { mm[imax, j]=vsolming[j];}
                mm[imax, 0] = foming;               
                goto atualizafoming;                
            }

            if (fominant == fomin) { contr++; }
            else { contr = 0; }


            if (contr>1000)
            {
                for (var i = 1; i <= npop; i++)
                {
                    for (var j = 1; j <= 580; j++)
                    { mm[i, j] = aleat.NextDouble() * (vlmax[j] - vlmin[j]) + vlmin[j]; }
                }
                for (var j = 1; j <= 580; j++) { mm[imax, j] = vsolming[j]; }
                mm[imax, 0] = foming;
                contr = 0;
                Console.WriteLine("Resetando a populacao mantendo a melhor solucao...");

                goto atualizafoming;
            }

            /* fazendo uma estratégia denominada de reset, se a diferença entre os valores
             * da função objetivo max e min ficar menor do que 1, significa que a população
             * está convergindo em um ótimo local. Portanto, se isto se repetir, seguidamente, 
             * por 50 vezes, será resetada a população ... geradas de forma aleatória e a melhor
             * solução até então será inserida em alguma posição (de forma aleatória tb).*/

            //escrever aqui...
            // Console.WriteLine(ite.ToString(), " min/max ", fomin.ToString(), " / ", fomax.ToString(), " ", contr.ToString());
            //Console.WriteLine("{0} min/max {1}/{2} - {3} - {4}", ite.ToString(), fomin.ToString(), fomax.ToString(), contr.ToString(),foming.ToString());
            //Console.WriteLine("{0} min/max {1}/{2} - {3}", ite.ToString(), fomin.ToString(), fomax.ToString(), contr.ToString());
            Console.WriteLine("{0} min/max {1}/{2} - {3}", ite.ToString(), fomin.ToString(), fomax.ToString(), ncorte.ToString());
            //texto retirado 1

            //seleção do tipo torneio com n=2

            for (var i = 1; i <= npop; i++)
            {
                aleat1 = aleat.Next(1, npop);
                aleat2 = aleat.Next(1, npop);
                if (mm[aleat1, 0] < mm[aleat2, 0]) { for (var j = 1; j <= 580; j++) { ms[i, j] = mm[aleat1, j]; } }
                else { for (var j = 1; j <= 580; j++) { ms[i, j] = mm[aleat2, j]; } }                
            }

        }

        static void populacao_inicial()
        {
            for (var i = 1; i <= npop; i++)
            {
                for (var j = 1; j <= 580; j++)
                { mm[i, j] = aleat.NextDouble() * (vlmax[j] - vlmin[j]) + vlmin[j]; }
            }
        }
       
        static void cruzamento()
        {
           /*cruzamento do tipo BLX-alpha , alpha adotado 0.5 
            * segundo a literatura esse é o melhor tipo de cruzamento para 
            * variáveis reais. Existe um parâmetro beta que pode ser gerado por
            * indivíduo ou por gene, optei por gene para ficar mais diversificado, 
            * tendo em vista que os  das variações por variáveis são diferentes.
            * Adotei cruzamento em 100% dos indivíduos.
            * Caso um gene ultrapasse os limites, em vez de refazer o cruzamento (pois 
            * isto deixaria o algoritmo mais lento, coloquei o valor do próprio limite 
            * caso isto aconteca). */
            double genep1,genep2,genef1,genef2,beta; //aleatorios reais
            //= 2 * ALEATÓRIO() - 0,5
            for (var i = 1; i <= npop/2; i++)
            {
                for (var j = 1; j <= 580; j++)
                {
                    genep1 = ms[2 * i - 1, j];
                    genep2 = ms[2 * i, j];
                    beta = (1 + 2 * alpha) * aleat.NextDouble() - alpha;
                    genef1 = beta * genep1 + (1 - beta) * genep2;
                    genef2 = (1 - beta) * genep1 + beta * genep2;
                    //verificando os limites
                    if (genef1 < vlmin[j]) { genef1 = vlmin[j]; }
                    if (genef1 > vlmax[j]) { genef1 = vlmax[j]; }
                    if (genef2 < vlmin[j]) { genef2 = vlmin[j]; }
                    if (genef2 > vlmax[j]) { genef2 = vlmax[j]; }
                    //jogando na matriz final do cruzamento mc
                    mc[2 * i - 1, j] = genef1;
                    mc[2 * i, j] = genef2;
                }


            }
        }

        static void mutacao()
        {            
            for (var i = 1; i <= npop; i++)
            {
                for (var j = 1; j <= 580; j++)
                {
                    if (aleat.NextDouble() <= pmut)
                    {
                        mm[i, j] = aleat.NextDouble() * (vlmax[j] - vlmin[j]) + vlmin[j];
                    }
                    else
                    {
                        mm[i, j] = mc[i, j];
                    }
                }       
            }
        }

        static void ordenando()
        {
            mfinal = new double[listasd.Count+1, 581];
            int cont = 0;
            string st1 = "";
            foreach (DictionaryEntry solucao in listasd)
            {
                cont++;                
                mfinal[cont,0]=Convert.ToDouble(solucao.Key.ToString());                
                st1 = solucao.Value.ToString();
                string[] vst1 = st1.Split(';');
                for (var j = 0; j <= 579;j++) mfinal[cont,j+1] = Convert.ToDouble(vst1[j]); 
            }
            mfinal[0, 0] = cont;
            double[] vaux1 = new double[581];
            bool fim = false;
            int i= 0;
            int k=0;
            while (fim == false)
            {
                i++;               
                if (mfinal[i+1,0]<mfinal[i,0])
                {                    
                    for (var j = 0; j <= 580; j++) vaux1[j] = mfinal[i, j];
                    for (var j = 0; j <= 580; j++) mfinal[i, j] = mfinal[i + 1, j];
                    for (var j = 0; j <= 580; j++) mfinal[i + 1, j] = vaux1[j];
                    k++;
                }
                if (i == listasd.Count-1)
                {
                    i = 0;
                    if (k == 0) fim = true;
                    if (k > 0) k = 0;
                    
                }                                    
            }
        }

        static void salvando()
        {
            string path = Environment.CurrentDirectory;
            path += "\\melhores_solucoes.xlsx";
            ExcelPackage arquivoexcel = new ExcelPackage(new FileInfo(@path));
            ExcelWorkbook planilha = arquivoexcel.Workbook;
            ExcelWorksheet aba;

            string st1 = DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString()+"("+corte.ToString()+")";
            planilha.Worksheets.Add(st1);
            planilha.Worksheets.MoveToStart(st1);
            aba = planilha.Worksheets[1];
            for (var j = 1; j <= 582; j++) aba.Column(j).Width = 20;
            for (var j = 3; j <= 582; j++) aba.Cells[1, j].Value = "param " + (j - 2).ToString();
            aba.Cells[1, 1].Value = "Soluções";
            aba.Cells[1, 2].Value = "F obj";
            for (var j = 1; j <= 581; j++) aba.Cells[1, j].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            int lisup = Convert.ToInt32(mfinal[0, 0]);
            if (lisup > 100) lisup = 100;
            for (var i = 1; i <= lisup; i++)
            {
                aba.Cells[i + 1, 1].Value = i;
                for (var j =0; j <= 580; j++) aba.Cells[i+1, j + 2].Value = mfinal[i, j];
            }
                
                

                
                   
            arquivoexcel.Save();
        }
        
    }
}

